VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Class Modules: XMLBuilder
'
' Description:  XMLBuilder provides a set of basic methods to
'               build a well-formed XML document.  These methods
'               take into account special characters and are
'               able to convert them to their corresponding
'               XML representations.  This class uses string
'               concatenation to build the XML document.  An
'               alternative way to build an XML document using
'               a DOM parser is shown in DOMXMLBuilder class
'               (see ModifyCustomer sample).
'
'               This class may throw error.  The caller of this class
'               must have exception handling setup (On Error GoTo ...)
'
' Created On: 06/15/2001
' Updated to SDK 2.0 On: 07/30/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'----------------------------------------------------------

Option Explicit

' XML string
Private m_XML  As String

' XML indentation
Private Const c_IndentationChars As String = "    " ' 4 spaces
Private m_IndentLevel     As Integer


' Return the XML String
Public Function GetXML() As String
    GetXML = m_XML
End Function

' Initialize members
Public Sub Init()
    
    m_XML = ""
    m_IndentLevel = 0
        
End Sub


Public Sub AddTagTopLevel(inputStr As String)
    
    m_XML = m_XML & "<" & inputStr & ">" & vbCrLf
    
End Sub

' Add aggregate with no attribute
Public Sub AddAggregateTagOnly(inputTag As String)
                             
    AddIndentation
    
    m_XML = m_XML & "<" & inputTag
    m_XML = m_XML & ">" & vbCrLf
    
    m_IndentLevel = m_IndentLevel + 1

End Sub

' Add aggregate with attributes
Public Sub AddAggregate(inputTag As String, _
                        attNameArray() As String, _
                        attValueArray() As String, _
                        attCount As Integer)
                                                
    AddIndentation
    
    m_XML = m_XML & "<" & inputTag
    
    Dim i As Integer
    Dim replacedVal As String
    
    ' put attribute list there
    For i = 0 To attCount - 1
        
        ' handle predefined chars
        If HasPreDefinedChar(attValueArray(i)) Then
            replacedVal = ReplacePreDefinedChar(attValueArray(i))
        Else
            replacedVal = attValueArray(i)
        End If
        m_XML = m_XML & " " & attNameArray(i) & _
                     "=" & """" & replacedVal & """" & " "
    Next
    
    ' end of tag
    m_XML = m_XML & ">" & vbCrLf
    
    m_IndentLevel = m_IndentLevel + 1
    
    Exit Sub
    
End Sub

Public Sub AddEndAggregate(inputTag As String)
    
    ' shift to left
    m_IndentLevel = m_IndentLevel - 1
    
    AddIndentation
    
    m_XML = m_XML & "</" & inputTag & ">" & vbCrLf
    
End Sub

' Add an element
Public Sub AddElement(inTag As String, value As String)
        
    AddIndentation
    
    ' handle predefined chars
    Dim replacedVal As String
    If HasPreDefinedChar(value) Then
        replacedVal = ReplacePreDefinedChar(value)
    Else
        replacedVal = value
    End If
    
    m_XML = m_XML & "<" & _
            inTag & ">" & _
            replacedVal & _
            "</" & inTag & _
            ">" & _
            vbCrLf
                 
End Sub

Private Sub AddIndentation()
    Dim i As Integer
    
    ' write space
    For i = 1 To m_IndentLevel
        m_XML = m_XML & c_IndentationChars
    Next
    
End Sub

Private Function HasPreDefinedChar(inputStr As String) As Boolean
    HasPreDefinedChar = False
    
    Dim result As Integer
    result = InStr(1, inputStr, "<", vbTextCompare)
    If Not (result = 0) Then
        HasPreDefinedChar = True
        Exit Function
    End If
    
    result = InStr(1, inputStr, ">", vbTextCompare)
    If Not (result = 0) Then
        HasPreDefinedChar = True
        Exit Function
    End If
    
    result = InStr(1, inputStr, "&", vbTextCompare)
    If Not (result = 0) Then
        HasPreDefinedChar = True
        Exit Function
    End If
    
    result = InStr(1, inputStr, "'", vbTextCompare)
    If Not (result = 0) Then
        HasPreDefinedChar = True
        Exit Function
    End If
    
    result = InStr(1, inputStr, """", vbTextCompare)
    If Not (result = 0) Then
        HasPreDefinedChar = True
        Exit Function
    End If
        
    HasPreDefinedChar = False
    
End Function

Private Function ReplacePreDefinedChar(inputStr As String) As String
    Dim replacedStr As String
   
    If HasPreDefinedChar(inputStr) Then
        replacedStr = Replace(inputStr, "&", "&amp;")
        replacedStr = Replace(replacedStr, "<", "&lt;")
        replacedStr = Replace(replacedStr, ">", "&gt;")
        replacedStr = Replace(replacedStr, "'", "&apos;")
        replacedStr = Replace(replacedStr, """", "&quot;")
        ReplacePreDefinedChar = replacedStr
    Else
        ReplacePreDefinedChar = inputStr
    End If

End Function




