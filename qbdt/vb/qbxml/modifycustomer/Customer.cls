VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Class Module: Customer
'
' Description:  Provides get/let methods to access customer data.
'
' Created On: 11/09/2001
' Updated to SDK 2.0: 08/05/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Private m_ListID As String
Private m_TimeCreated As String
Private m_TimeModified As String
Private m_EditSequence As String
Private m_Name As String
Private m_FullName As String
Private m_IsActive As String
Private m_Sublevel As String
Private m_FirstName As String
Private m_LastName As String
Private m_CompanyName As String
Private m_BillAddress_Addr1 As String
Private m_BillAddress_Addr2 As String
Private m_BillAddress_City As String
Private m_BillAddress_State As String
Private m_BillAddress_PostalCode As String
Private m_Phone As String
Private m_Email As String
Private m_Contact As String
 
Private m_CustomerTypeRef_ListID As String
Private m_CustomerTypeRef_Residential As String
      
Private m_TermsRef_ListID As String
Private m_TermsRef_FullName As String

Private m_Balance As String
Private m_TotalBalance As String

Private m_SalesTaxCodeRef_ListID As String
Private m_SalesTaxCodeRef_FullName As String

Private m_ItemSalesTaxRef_ListID As String
Private m_ItemSalesTaxRef_FullName As String
 
Private m_Status As String

Public Property Get ListID() As String
    ListID = m_ListID
End Property

Public Property Let ListID(ByVal value As String)
    m_ListID = value
End Property

Public Property Get FullName() As String
    FullName = m_FullName
End Property

Public Property Let FullName(ByVal value As String)
    m_FullName = value
End Property

Public Property Get TimeCreated() As String
    TimeCreated = m_TimeCreated
End Property

Public Property Let TimeCreated(ByVal value As String)
    m_TimeCreated = value
End Property

Public Property Get TimeModified() As String
    TimeModified = m_TimeModified
End Property

Public Property Let TimeModified(ByVal value As String)
    m_TimeModified = value
End Property

Public Property Get EditSequence() As String
    EditSequence = m_EditSequence
End Property

Public Property Let EditSequence(ByVal value As String)
    m_EditSequence = value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal value As String)
    m_Name = value
End Property

Public Property Get IsActive() As String
    IsActive = m_IsActive
End Property

Public Property Let IsActive(ByVal value As String)
    m_IsActive = value
End Property

Public Property Get Sublevel() As String
    Sublevel = m_Sublevel
End Property

Public Property Let Sublevel(ByVal value As String)
    m_Sublevel = value
End Property

Public Property Get FirstName() As String
    FirstName = m_FirstName
End Property

Public Property Let FirstName(ByVal value As String)
    m_FirstName = value
End Property

Public Property Get LastName() As String
    LastName = m_LastName
End Property

Public Property Let LastName(ByVal value As String)
    m_LastName = value
End Property

Public Property Get BillAddress_Addr1() As String
    BillAddress_Addr1 = m_BillAddress_Addr1
End Property

Public Property Let BillAddress_Addr1(ByVal value As String)
    m_BillAddress_Addr1 = value
End Property

Public Property Get BillAddress_Addr2() As String
    BillAddress_Addr2 = m_BillAddress_Addr2
End Property

Public Property Let BillAddress_Addr2(ByVal value As String)
    m_BillAddress_Addr2 = value
End Property

Public Property Get BillAddress_City() As String
    BillAddress_City = m_BillAddress_City
End Property

Public Property Let BillAddress_City(ByVal value As String)
    m_BillAddress_City = value
End Property

Public Property Get BillAddress_State() As String
    BillAddress_State = m_BillAddress_State
End Property

Public Property Let BillAddress_State(ByVal value As String)
    m_BillAddress_State = value
End Property

Public Property Get BillAddress_PostalCode() As String
    BillAddress_PostalCode = m_BillAddress_PostalCode
End Property

Public Property Let BillAddress_PostalCode(ByVal value As String)
    m_BillAddress_PostalCode = value
End Property

Public Property Get Phone() As String
    Phone = m_Phone
End Property

Public Property Let Phone(ByVal value As String)
    m_Phone = value
End Property

Public Property Get Email() As String
    Email = m_Email
End Property

Public Property Let Email(ByVal value As String)
    m_Email = value
End Property

Public Property Get Contact() As String
    Contact = m_Contact
End Property

Public Property Let Contact(ByVal value As String)
    m_Contact = value
End Property

Public Property Get CustomerTypeRef_ListID() As String
    CustomerTypeRef_ListID = m_CustomerTypeRef_ListID
End Property

Public Property Let CustomerTypeRef_ListID(ByVal value As String)
    m_CustomerTypeRef_ListID = value
End Property

Public Property Get CustomerTypeRef_Residential() As String
    CustomerTypeRef_Residential = m_CustomerTypeRef_Residential
End Property

Public Property Let CustomerTypeRef_Residential(ByVal value As String)
    m_CustomerTypeRef_Residential = value
End Property

Public Property Get TermsRef_ListID() As String
    TermsRef_ListID = m_TermsRef_ListID
End Property

Public Property Let TermsRef_ListID(ByVal value As String)
    m_TermsRef_ListID = value
End Property

Public Property Get TermsRef_FullName() As String
    TermsRef_FullName = m_TermsRef_FullName
End Property

Public Property Let TermsRef_FullName(ByVal value As String)
    m_TermsRef_FullName = value
End Property

Public Property Get CompanyName() As String
    CompanyName = m_CompanyName
End Property

Public Property Let CompanyName(ByVal value As String)
    m_CompanyName = value
End Property

Public Property Get Balance() As String
    Balance = m_Balance
End Property

Public Property Let Balance(ByVal value As String)
    m_Balance = value
End Property

Public Property Get TotalBalance() As String
    TotalBalance = m_TotalBalance
End Property

Public Property Let TotalBalance(ByVal value As String)
    m_TotalBalance = value
End Property

Public Property Get SalesTaxCodeRef_ListID() As String
    SalesTaxCodeRef_ListID = m_SalesTaxCodeRef_ListID
End Property

Public Property Let SalesTaxCodeRef_ListID(ByVal value As String)
    m_SalesTaxCodeRef_ListID = value
End Property

Public Property Get SalesTaxCodeRef_FullName() As String
    SalesTaxCodeRef_FullName = m_SalesTaxCodeRef_FullName
End Property

Public Property Let SalesTaxCodeRef_FullName(ByVal value As String)
    m_SalesTaxCodeRef_FullName = value
End Property

Public Property Get ItemSalesTaxRef_ListID() As String
    ItemSalesTaxRef_ListID = m_ItemSalesTaxRef_ListID
End Property

Public Property Let ItemSalesTaxRef_ListID(ByVal value As String)
    m_ItemSalesTaxRef_ListID = value
End Property

Public Property Get ItemSalesTaxRef_fullName() As String
    ItemSalesTaxRef_fullName = m_ItemSalesTaxRef_FullName
End Property

Public Property Let ItemSalesTaxRef_fullName(ByVal value As String)
    m_ItemSalesTaxRef_FullName = value
End Property


