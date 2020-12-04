VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InfoClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------
' Class: InfoClass
'
' Description: This class contains RefNumber and TxnID Strings.
'           The properties will set and get the RefNumber or TxnID.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

    Private m_sRefNumber As String
    Private m_sTxnID As String

    Function SetTxnID(ByVal Value As String)
        m_sTxnID = Value
    End Function
    Function GetTxnID() As String
        GetTxnID = m_sTxnID
    End Function
    Function SetRefNumber(ByVal Value As String)
        m_sRefNumber = Value
    End Function

    Function GetRefNumber() As String
        GetRefNumber = m_sRefNumber
    End Function
    
