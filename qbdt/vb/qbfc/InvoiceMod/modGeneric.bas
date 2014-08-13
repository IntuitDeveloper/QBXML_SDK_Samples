Attribute VB_Name = "modGeneric"
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'----------------------------------------------------------

Option Explicit

Dim Implementation As String
Dim strInvoiceLineInfo As String

Dim strqbXMLLevel As String
Dim booSupportsModify As Boolean
Dim booSupportsDateTime As Boolean


Public Sub SetImplementation(Value As String)
  Implementation = Value
End Sub


Public Function GetImplementation() As String
  GetImplementation = Implementation
End Function


Public Function SupportsModify() As Boolean
  SupportsModify = booSupportsModify
End Function


Public Function SupportsDateTime() As Boolean
  SupportsDateTime = booSupportsDateTime
End Function


Public Sub SetInvoiceLineInfo(strInvoiceLineInfoIn As String)
  strInvoiceLineInfo = strInvoiceLineInfoIn
End Sub


Public Function GetInvoiceLineInfo() As String
  GetInvoiceLineInfo = strInvoiceLineInfo
End Function


Public Function QuickBooksVersionOK() As Boolean
  'The OpenConnectionBeginSession routines will exit this sample program
  'if installation or access problems are found
  If Implementation = "QBXMLRP" Then
    QBXMLRP_OpenConnectionBeginSession
  Else
    QBFC_OpenConnectionBeginSession
  End If
  
  'Now check to see what version of QuickBooks is running and what version
  'of qbXML it supports
  strqbXMLLevel = GetMaxVersionSupported
  
  'Check for version 2.1 or 3.X
  If Not (InStr(1, strqbXMLLevel, "2.1") > 0) And Not (InStr(1, strqbXMLLevel, "3.") > 0) And Not (InStr(1, strqbXMLLevel, "4.") > 0) And Not (InStr(1, strqbXMLLevel, "5.") > 0) Then
    MsgBox "The configuration you're running against " & _
      "does not support Invoice Modify." & vbCrLf & vbCrLf & _
      "You will be able to query " & _
      "for invoices, but you will not be able to modify them."
    booSupportsModify = False
  Else
    booSupportsModify = True
  End If
  
  If InStr(1, strqbXMLLevel, "2") Or InStr(1, strqbXMLLevel, "3") Then
    booSupportsDateTime = True
  Else
    booSupportsDateTime = False
  End If
  
  QuickBooksVersionOK = True
End Function


Public Sub EndSessionCloseConnection()
  If Implementation = "QBXMLRP" Then
    QBXMLRP_EndSessionCloseConnection
  ElseIf Implementation = "QBFC" Then
    QBFC_EndSessionCloseConnection
'  Else
'    QBFCCA_EndSessionCloseConnection
  End If
End Sub


Function GetMaxVersionSupported() As String
  If Implementation = "QBXMLRP" Then
    GetMaxVersionSupported = QBXMLRP_MaxVersionSupported
  ElseIf Implementation = "QBFC" Then
    GetMaxVersionSupported = QBFC_MaxVersionSupported
'  Else
'    GetMaxVersionSupported = QBFCCA_MaxVersionSupported
  End If
End Function


Public Sub FillComboBox(cmbComboBox As ComboBox, _
                        strQueryType As String, _
                        strNameElement As String, _
                        strFilter As String, _
                        booMarkGroupItems As Boolean)

  If Implementation = "QBXMLRP" Then
    QBXMLRP_FillComboBox _
      cmbComboBox, strQueryType, strNameElement, strFilter, booMarkGroupItems
  ElseIf Implementation = "QBFC" Then
    QBFC_FillComboBox _
      cmbComboBox, strQueryType, strNameElement, strFilter, booMarkGroupItems
'  Else
'    QBFCCA_FillComboBox _
'      cmbComboBox, strQueryType, strNameElement, strFilter, booMarkGroupItems
  End If
End Sub


Public Sub FillInvoiceList(lstInvoices As ListBox, _
                           strRefNumber As String, _
                           strFromDateTime As String, _
                           strToDateTime As String, _
                           strDateQueryType As String, _
                           strDateMacro As String, _
                           strCustomerJob As String, _
                           booCustomerWithChildren As Boolean, _
                           strAccount As String, _
                           booAccountWithChildren As Boolean, _
                           strFromRefNumberRange As String, _
                           strToRefNumberRange As String, _
                           strRefNumberPiece As String, _
                           strRefNumberCriteria As String, _
                           strPaidStatus As String)

  If Implementation = "QBXMLRP" Then
    QBXMLRP_FillInvoiceList lstInvoices, strRefNumber, strFromDateTime, _
      strToDateTime, strDateQueryType, strDateMacro, strCustomerJob, _
      booCustomerWithChildren, strAccount, booAccountWithChildren, _
      strFromRefNumberRange, strToRefNumberRange, strRefNumberPiece, _
      strRefNumberCriteria, strPaidStatus
  ElseIf Implementation = "QBFC" Then
    QBFC_FillInvoiceList lstInvoices, strRefNumber, strFromDateTime, _
      strToDateTime, strDateQueryType, strDateMacro, strCustomerJob, _
      booCustomerWithChildren, strAccount, booAccountWithChildren, _
      strFromRefNumberRange, strToRefNumberRange, strRefNumberPiece, _
      strRefNumberCriteria, strPaidStatus
'  Else
'    QBFCCA_FillInvoiceList lstInvoices, strRefNumber, strFromDateTime, _
'      strToDateTime, strDateQueryType, strDateMacro, strCustomerJob, _
'      booCustomerWithChildren, strAccount, booAccountWithChildren, _
'      strFromRefNumberRange, strToRefNumberRange, strRefNumberPiece, _
'      strRefNumberCriteria, strPaidStatus
  End If
End Sub


Public Sub GetInvoice(TxnID As String)
  If Implementation = "QBXMLRP" Then
    QBXMLRP_GetInvoice TxnID
  ElseIf Implementation = "QBFC" Then
    QBFC_GetInvoice TxnID
'  Else
'    QBFCCA_GetInvoice TxnID
  End If
End Sub


Public Sub LoadInvoiceModifyForm()
  If Implementation = "QBXMLRP" Then
    QBXMLRP_LoadInvoiceModifyForm
  ElseIf Implementation = "QBFC" Then
    QBFC_LoadInvoiceModifyForm
'  Else
'    QBFCCA_LoadInvoiceModifyForm
  End If
End Sub


Public Sub LoadInvoiceLineArray(strLineArray() As String)
  If Implementation = "QBXMLRP" Then
    QBXMLRP_LoadInvoiceLineArray strLineArray
  ElseIf Implementation = "QBFC" Then
    QBFC_LoadInvoiceLineArray strLineArray
'  Else
'    QBFCCA_LoadInvoiceLineArray strLineArray
  End If
End Sub


Public Sub ModifyInvoice(strInvoiceChangeString As String)
  If Implementation = "QBXMLRP" Then
    QBXMLRP_ModifyInvoice strInvoiceChangeString
  ElseIf Implementation = "QBFC" Then
    QBFC_ModifyInvoice strInvoiceChangeString
'  Else
'    QBFCCA_ModifyInvoice strInvoiceChangeString
  End If
End Sub


Public Sub GetItemInfo(strItemFullName As String, _
                       strDesc As String, _
                       strRate As String, _
                       strRateOrPercent As String, _
                       strSalesTaxCode As String)

  strDesc = Empty
  strRate = Empty
  strRateOrPercent = "Rate"
  strSalesTaxCode = Empty

  If Implementation = "QBXMLRP" Then
    QBXMLRP_GetItemInfo _
      strItemFullName, strDesc, strRate, strRateOrPercent, strSalesTaxCode
  ElseIf Implementation = "QBFC" Then
    QBFC_GetItemInfo _
      strItemFullName, strDesc, strRate, strRateOrPercent, strSalesTaxCode
'  Else
'    QBFCCA_GetItemInfo _
'      strItemFullName, strDesc, strRate, strRateOrPercent, strSalesTaxCode
  End If
End Sub


Public Function RequestInfo() As String

  Dim strRequestText As String
  
  If Implementation = "QBXMLRP" Then
    QBXMLRP_LoadRequest strRequestText
  ElseIf Implementation = "QBFC" Then
    QBFC_LoadRequest strRequestText
'  Else
'    QBFCCA_LoadRequest strRequestText
  End If
  
  RequestInfo = strRequestText
End Function
