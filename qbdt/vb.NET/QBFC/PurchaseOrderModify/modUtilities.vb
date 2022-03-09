Option Strict Off
Option Explicit On
Module modUtilities
	'----------------------------------------------------------
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Dim stroldMessageSetID As String

    Public Sub AddMSXMLNode(ByRef strNodeName As String, ByRef objDOMDocument As MSXML2.DOMDocument60, ByRef objParentNode As MSXML2.IXMLDOMNode, ByRef objNode As MSXML2.IXMLDOMNode)

        objNode = objDOMDocument.createElement(strNodeName)
        objParentNode.appendChild(objNode)
    End Sub


    Public Sub AddMSXMLAttribute(ByRef strAttributeName As String, ByRef strAttributeValue As String, ByRef objDOMDocument As MSXML2.DOMDocument60, ByRef objParentNode As MSXML2.IXMLDOMNode, ByRef objAttribute As MSXML2.IXMLDOMAttribute)

        objAttribute = objDOMDocument.createAttribute(strAttributeName)
        objAttribute.text = strAttributeValue
        
        objParentNode.attributes.setNamedItem(objAttribute)
    End Sub

    Public Sub AddMSXMLElement(ByRef strElementName As String, ByRef strElementValue As String, ByRef objDOMDocument As MSXML2.DOMDocument60, ByRef objParentNode As MSXML2.IXMLDOMNode, ByRef objElement As MSXML2.IXMLDOMElement)


        objElement = objDOMDocument.createElement(strElementName)
        objElement.text = strElementValue
        
        objParentNode.appendChild(objElement)
    End Sub


    Public Sub CreateStandardRequestNode(ByRef boonewMessageSetID As Boolean, ByRef stronError As String, ByRef objDOMDocument As MSXML2.DOMDocument60, ByRef objRootNode As MSXML2.IXMLDOMNode, ByRef objRequestNode As MSXML2.IXMLDOMNode, ByRef objAttribute As MSXML2.IXMLDOMAttribute)

        objRootNode = objDOMDocument.createElement("QBXML")
        objDOMDocument.appendChild(objRootNode)

        objRequestNode = objDOMDocument.createElement("QBXMLMsgsRq")
        objRootNode.appendChild(objRequestNode)

        
        If Not IsNothing(stroldMessageSetID) Then
            AddMSXMLAttribute("oldMessageSetID", stroldMessageSetID, objDOMDocument, objRequestNode, objAttribute)
        End If

        stroldMessageSetID = CStr(Nothing)

        If boonewMessageSetID Then
            stroldMessageSetID = NewMessageSetID()
            AddMSXMLAttribute("newMessageSetID", stroldMessageSetID, objDOMDocument, objRequestNode, objAttribute)
        End If

        AddMSXMLAttribute("onError", stronError, objDOMDocument, objRequestNode, objAttribute)
    End Sub


    Private Function NewMessageSetID() As String
		'Because this is a sample program that isn't expected to run on multiple
		'machines and isn't expected to produce more than one modify request per
		'second, we can simply use the date time string for our message set ID.
		'If we were expecting multiple machines we'd add some sort of system
		'identifier.  If we were expecting more than one request message set per
		'second we'd use a smaller time sequence.
		NewMessageSetID = VB6.Format(Now, "yyyymmddThhmmss")
	End Function
	
	
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
		
		IndentString = CStr(Nothing)
		
		FileNum = FreeFile
		FileOpen(FileNum, XMLFile, OpenMode.Output)
		
		'Remove the linefeeds from the XML output string
		XMLString = Replace(XMLString, vbLf, vbNullString)
		
		SplitXMLString = Split(XMLString, "<")
		
		'We're expecting the first character of the XML output to be "<"
		'which result in an empty first array element, so skip it.
		
		SplitIndex = LBound(SplitXMLString) + 1
		
		Do 
			
			If Left(SplitXMLString(SplitIndex), 1) = "/" Then
				IndentString = Left(IndentString, Len(IndentString) - 3)
				
				PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
				
				SplitIndex = SplitIndex + 1
				
			ElseIf Left(SplitXMLString(SplitIndex + 1), 1) = "/" Then 
				
				If InStr(1, Left(SplitXMLString(SplitIndex), InStr(1, SplitXMLString(SplitIndex), ">")), " ") > 0 Then
					
					PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
					
					SplitIndex = SplitIndex + 1
				Else
					
					PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex) & "<" & SplitXMLString(SplitIndex + 1))
					
					SplitIndex = SplitIndex + 2
				End If
			Else
				
				PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
				
				If SplitXMLString(SplitIndex) <> "?xml version=""1.0"" ?>" And SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' >" And Right(SplitXMLString(SplitIndex), 2) <> "/>" Then
					IndentString = IndentString & "   "
				End If
				
				SplitIndex = SplitIndex + 1
			End If
			
		Loop Until SplitIndex >= UBound(SplitXMLString)
		
		If Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
			If Len(IndentString) >= 3 Then
				IndentString = Left(IndentString, Len(IndentString) - 3)
			End If
		End If
		
		PrintLine(FileNum, IndentString & "<" & SplitXMLString(UBound(SplitXMLString)))
		
		FileClose(FileNum)
	End Sub
End Module