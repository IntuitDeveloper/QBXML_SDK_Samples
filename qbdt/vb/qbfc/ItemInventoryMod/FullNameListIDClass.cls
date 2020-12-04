VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FullNameListIDClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------
' Class: FullNameListIDClass
'
' Description: This class contains FullName and ListID Strings.
'           The properties will set and get the ListID or FullName.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
    Private m_sFullName As String
    Private m_sListID As String

    Public Property Get ListID() As String
        ListID = m_sListID
    End Property
    
    Public Property Let ListID(value As String)
        m_sListID = value
    End Property
        
    Public Property Get FullName() As String
        FullName = m_sFullName
    End Property
    
    Public Property Let FullName(value As String)
        m_sFullName = value
    End Property

    Public Function ToString() As String
        ToString = m_sFullName
    End Function


