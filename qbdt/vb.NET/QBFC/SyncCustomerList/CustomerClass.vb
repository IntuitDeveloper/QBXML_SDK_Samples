'----------------------------------------------------------
' Class: CustomerClass
'
' Description: This class contains customerRet object.
'           The properties will return the ListID or FullName
'           or get/set the customerRet object.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
Imports Interop.QBFC13

Public Class CustomerClass
    Inherits Object

    Private m_customerRet As ICustomerRet

    ReadOnly Property ListID() As String
        Get
            Return m_customerRet.ListID.GetValue
        End Get
    End Property   ' ListID

    ReadOnly Property FullName() As String
        Get
            Return m_customerRet.FullName.GetValue
        End Get
    End Property   ' FullName

    Property customerRet() As ICustomerRet
        Get
            Return m_customerRet
        End Get
        Set(ByVal Value As ICustomerRet)
            m_customerRet = Value
        End Set
    End Property   ' customerRet

    Public Overrides Function ToString() As String
        ToString = m_customerRet.FullName.GetValue()
    End Function

End Class
