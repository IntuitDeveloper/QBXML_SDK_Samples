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
Public Class FullNameListIDClass
    Inherits Object

    Private m_sFullName As String
    Private m_sListID As String

    Property ListID() As String
        Get
            Return m_sListID
        End Get
        Set(ByVal Value As String)
            m_sListID = Value
        End Set
    End Property   ' ListID

    Property FullName() As String
        Get
            Return m_sFullName
        End Get
        Set(ByVal Value As String)
            m_sFullName = Value
        End Set
    End Property   ' FullName

    Public Overrides Function ToString() As String
        ToString = m_sFullName
    End Function

End Class
