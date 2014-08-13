VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    Private m_customerRet As ICustomerRet

    Public Function GetListID() As String
        If (m_customerRet Is Nothing) Then
            GetListID = ""
        Else
            GetListID = m_customerRet.listID.GetValue
        End If
    End Function
    
    Public Function GetFullName() As String
        If (m_customerRet Is Nothing) Then
            GetFullName = ""
        Else
            GetFullName = m_customerRet.FullName.GetValue
        End If
    End Function
    
    Public Function GetCustomerRet() As ICustomerRet
            GetCustomerRet = m_customerRet
    End Function
    
    Public Function SetCustomerRet(ByRef Value As ICustomerRet)
            Set m_customerRet = Value
    End Function

    Public Function ToString() As String
        ToString = m_sFullName
    End Function
