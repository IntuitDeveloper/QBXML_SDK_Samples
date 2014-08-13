Attribute VB_Name = "modGeneric"
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Dim Implementation As String

Dim strqbXMLLevel As String
Dim booSupportsModify As Boolean

Dim dateFrom As Date
Dim dateTo As Date

Dim strSelectedPOInfo As String


Public Sub SetImplementation(Value As String)
  Implementation = Value
End Sub


Public Function GetImplementation() As String
  GetImplementation = Implementation
End Function


Public Sub SetDates(dateInFrom As Date, _
                    dateInTo As Date)

  dateFrom = dateInFrom
  dateTo = dateInTo
End Sub


Public Sub SetSelectedPOInfo(strPOInfo As String)
  strSelectedPOInfo = strPOInfo
End Sub


Public Function SupportsModify() As Boolean
  SupportsModify = booSupportsModify
End Function

Public Function QuickBooksVersionOK() As Boolean
  'This used to show error checking, but it did so badly, so we just
  'made a dummy out of this function.  If you run this against a QB
  'version less than QB2003 R3 you'll get a runtime error.
  If Implementation = "QBXMLRP" Then
    QBXMLRP_OpenConnectionBeginSession
  Else
    QBFC_OpenConnectionBeginSession
  End If
  
  strqbXMLLevel = GetMaxVersionSupported
  
    booSupportsModify = True
  
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
    GetMaxVersionSupported = "3.0"
  ElseIf Implementation = "QBFC" Then
    GetMaxVersionSupported = "3.0"
  End If
End Function

Public Sub LoadPOInfoArray(strPOInfo() As String, _
                           booGiveWarning As Boolean)
  If Implementation = "QBXMLRP" Then
    QBXMLRP_LoadPOInfoArray strPOInfo, dateFrom, dateTo, booGiveWarning
  ElseIf Implementation = "QBFC" Then
    QBFC_LoadPOInfoArray strPOInfo, dateFrom, dateTo, booGiveWarning
'  Else
'    QBFCCA_LoadPOInfoArray strPOInfo, dateFrom, dateTo, booGiveWarning
  End If
End Sub

Public Sub GetPOInfo(strTxnID As String, _
                     strEditSequence As String, _
                     strRefNumber As String, _
                     strTxnDate As String, _
                     strVendor As String, _
                     strPOLines() As String)
  
  If Implementation = "QBXMLRP" Then
    QBXMLRP_GetPOInfo strTxnID, strEditSequence, strRefNumber, strTxnDate, _
                      strVendor, strPOLines, strSelectedPOInfo
  ElseIf Implementation = "QBFC" Then
    QBFC_GetPOInfo strTxnID, strEditSequence, strRefNumber, strTxnDate, _
                   strVendor, strPOLines, strSelectedPOInfo
'  Else
'    QBFCCA_GetPOInfo strTxnID, strEditSequence, strRefNumber, strTxnDate, _
'                     strVendor, strPOLines, strSelectedPOInfo
  End If
End Sub


Public Sub ClosePO(strTxnID As String, _
                   strEditSequence As String)
  If Implementation = "QBXMLRP" Then
    QBXMLRP_ClosePO strTxnID, strEditSequence
  ElseIf Implementation = "QBFC" Then
    QBFC_ClosePO strTxnID, strEditSequence
'  Else
'    QBFCCA_ClosePO strTxnID, strEditSequence
  End If
End Sub


Public Sub ChangePOLine(strAction As String, _
                        strTxnID As String, _
                        strEditSequence As String, _
                        strPOLines() As String, _
                        intSelectedPOLine As Integer)

  Dim intGroupLine As Integer
  intGroupLine = -1
  If strAction <> "ReceiveAll" Then
    If LineType(strPOLines(intSelectedPOLine)) = "GroupSubItem" Then
      intGroupLine = intSelectedPOLine - 1
      Do While LineType(strPOLines(intGroupLine)) <> "GroupItem"
        intGroupLine = intGroupLine - 1
      Loop
    End If
  End If
  
  If Implementation = "QBXMLRP" Then
    QBXMLRP_ChangePOLine strAction, strTxnID, strEditSequence, _
                         strPOLines, intSelectedPOLine, intGroupLine
  ElseIf Implementation = "QBFC" Then
    QBFC_ChangePOLine strAction, strTxnID, strEditSequence, _
                      strPOLines, intSelectedPOLine, intGroupLine
'  Else
'    QBFCCA_ChangePOLine strAction, strTxnID, strEditSequence, _
'                        strPOLines, intSelectedPOLine, intGroupLine
  End If
End Sub


Public Sub SetQuantitiesAndBillForRemainingItems( _
             strTxnID As String, _
             strEditSequence As String, _
             strVendor As String, _
             strRefNumber As String, _
             strTxnDate As String, _
             strPOLines() As String, _
             intPOLine As Integer)

  If Implementation = "QBXMLRP" Then
    QBXMLRP_SetQuantitiesAndBillForRemainingItems _
      strTxnID, strEditSequence, strVendor, strRefNumber, _
      strTxnDate, strPOLines, intPOLine
  ElseIf Implementation = "QBFC" Then
    QBFC_SetQuantitiesAndBillForRemainingItems _
      strTxnID, strEditSequence, strVendor, strRefNumber, _
      strTxnDate, strPOLines, intPOLine
'  Else
'    QBFCCA_SetQuantitiesAndBillForRemainingItems _
'      strTxnID, strEditSequence, strVendor, strRefNumber, _
'      strTxnDate, strPOLines, intPOLine
  End If
End Sub


Public Function LineType(strPOLine As String) As String

  Dim Splits() As String
  
  Splits = Split(strPOLine, "<spliter>")
  
  LineType = Splits(9)
End Function
