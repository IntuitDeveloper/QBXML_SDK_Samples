Option Strict Off
Option Explicit On
Friend Class InfoClass
	'----------------------------------------------------------
	' Class: InfoClass
	'
	' Description: This class contains RefNumber and TxnID Strings.
	'           The properties will set and get the RefNumber or TxnID.
	'
	' Copyright © 2002-2020 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	Private m_sRefNumber As String
	Private m_sTxnID As String
	
	Function SetTxnID(ByVal Value As String) As Object
		m_sTxnID = Value
	End Function
	Function GetTxnID() As String
		GetTxnID = m_sTxnID
	End Function
	Function SetRefNumber(ByVal Value As String) As Object
		m_sRefNumber = Value
	End Function
	
	Function GetRefNumber() As String
		GetRefNumber = m_sRefNumber
	End Function
End Class