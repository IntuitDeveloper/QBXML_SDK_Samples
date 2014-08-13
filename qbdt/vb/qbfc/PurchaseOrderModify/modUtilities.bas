Attribute VB_Name = "modUtilities"
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Dim stroldMessageSetID As String

Public Sub AddMSXMLNode(strNodeName As String, _
                        objDOMDocument As DOMDocument, _
                        objParentNode As IXMLDOMNode, _
                        objNode As IXMLDOMNode)

  Set objNode = objDOMDocument.createElement(strNodeName)
  objParentNode.appendChild objNode
End Sub


Public Sub AddMSXMLAttribute(strAttributeName As String, _
                             strAttributeValue As String, _
                             objDOMDocument As DOMDocument, _
                             objParentNode As IXMLDOMNode, _
                             objAttribute As IXMLDOMAttribute)

  Set objAttribute = objDOMDocument.createAttribute(strAttributeName)
  objAttribute.Text = strAttributeValue
  objParentNode.Attributes.setNamedItem objAttribute
End Sub

Public Sub AddMSXMLElement(strElementName As String, _
                           strElementValue As String, _
                           objDOMDocument As DOMDocument, _
                           objParentNode As IXMLDOMNode, _
                           objElement As IXMLDOMElement)


  Set objElement = objDOMDocument.createElement(strElementName)
  objElement.Text = strElementValue
  objParentNode.appendChild objElement
End Sub


Public Sub CreateStandardRequestNode(boonewMessageSetID As Boolean, _
                                     stronError As String, _
                                     objDOMDocument As DOMDocument, _
                                     objRootNode As IXMLDOMNode, _
                                     objRequestNode As IXMLDOMNode, _
                                     objAttribute As IXMLDOMAttribute)

  Set objRootNode = objDOMDocument.createElement("QBXML")
  objDOMDocument.appendChild objRootNode
  
  Set objRequestNode = objDOMDocument.createElement("QBXMLMsgsRq")
  objRootNode.appendChild objRequestNode
  
  If stroldMessageSetID <> Empty Then
    AddMSXMLAttribute _
      "oldMessageSetID", stroldMessageSetID, objDOMDocument, objRequestNode, objAttribute
  End If
  
  stroldMessageSetID = Empty
  
  If boonewMessageSetID Then
    stroldMessageSetID = NewMessageSetID
    AddMSXMLAttribute _
      "newMessageSetID", stroldMessageSetID, objDOMDocument, objRequestNode, objAttribute
  End If
  
  AddMSXMLAttribute _
    "onError", stronError, objDOMDocument, objRequestNode, objAttribute
End Sub


Private Function NewMessageSetID() As String
  'Because this is a sample program that isn't expected to run on multiple
  'machines and isn't expected to produce more than one modify request per
  'second, we can simply use the date time string for our message set ID.
  'If we were expecting multiple machines we'd add some sort of system
  'identifier.  If we were expecting more than one request message set per
  'second we'd use a smaller time sequence.
  NewMessageSetID = Format(Now, "yyyymmddThhmmss")
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
Public Sub PrettyPrintXMLToFile(XMLString As String, _
                                 XMLFile As String)
                                       
  Dim SplitXMLString() As String
  Dim IndentString As String
  Dim XMLStringLength As Long
  Dim SplitIndex
  Dim FileNum As Integer
  
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
    If Len(IndentString) >= 3 Then
      IndentString = Left(IndentString, Len(IndentString) - 3)
    End If
  End If
  
  Print #FileNum, IndentString & "<" & _
                  SplitXMLString(UBound(SplitXMLString))
  
  Close FileNum
End Sub





