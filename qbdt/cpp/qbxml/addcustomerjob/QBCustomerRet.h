/*---------------------------------------------------------------------------
 * FILE: QBCustomerRet.h
 *
 * Description:
 * Provides get/set methods for CustomerRet object
 *
 * Created On: 11/11/2001
 * Updated to SDK 2.0: 08/08/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

//////////////////////////////////////////////////////////////////////

#if !defined(AFX_QBCUSTOMERRET_H__8B2D3111_B018_45FC_80C2_C9ABEA411886__INCLUDED_)
#define AFX_QBCUSTOMERRET_H__8B2D3111_B018_45FC_80C2_C9ABEA411886__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class QBCustomerRet
{
public:
  QBCustomerRet();
  ~QBCustomerRet();

  string &ListID() { return m_ListID;};
  void ListID(const string & value) { m_ListID = value;};

  string &TimeCreated() { return m_TimeCreated;};
  void TimeCreated(const string & value) { m_TimeCreated = value;};

  string &TimeModified() { return m_TimeModified;};
  void TimeModified(const string & value) { m_TimeModified = value;};

  string &EditSequence() { return m_EditSequence;};
  void EditSequence(const string & value) { m_EditSequence = value;};

  string &Name() { return m_Name;};
  void Name(const string & value) { m_Name = value;};

  string &FullName() { return m_FullName;};
  void FullName(const string & value) { m_FullName = value;};

  string &IsActive() { return m_IsActive;};
  void IsActive(const string & value) { m_IsActive = value;};

  string &ParentRef_ListID() { return m_ParentRef_ListID;};
  void ParentRef_ListID(const string & value) { m_ParentRef_ListID = value;};

  string &ParentRef_FullName() { return m_ParentRef_FullName;};
  void ParentRef_FullName(const string & value) { m_ParentRef_FullName = value;};

  string &Sublevel() { return m_Sublevel;};
  void Sublevel(const string & value) { m_Sublevel = value;};

  string &CompanyName() { return m_CompanyName;};
  void CompanyName(const string & value) { m_CompanyName = value;};

  string &FirstName() { return m_FirstName;};
  void FirstName(const string & value) { m_FirstName = value;};

  string &LastName() { return m_LastName;};
  void LastName(const string & value) { m_LastName = value;};

  string &BillAddress_Addr1() { return m_BillAddress_Addr1;};
  void BillAddress_Addr1(const string & value) { m_BillAddress_Addr1 = value;};

  string &BillAddress_Addr2() { return m_BillAddress_Addr2;};
  void BillAddress_Addr2(const string & value) { m_BillAddress_Addr2 = value;};

  string &BillAddress_City() { return m_BillAddress_City;};
  void BillAddress_City(const string & value) { m_BillAddress_City = value;};

  string &BillAddress_State() { return m_BillAddress_State;};
  void BillAddress_State(const string & value) { m_BillAddress_State = value;};

  string &BillAddress_PostalCode() { return m_BillAddress_PostalCode;};
  void BillAddress_PostalCode(const string & value) { m_BillAddress_State = value;};

  string &Phone() { return m_Phone;};
  void Phone(const string & value) { m_Phone = value;};

  string &Email() { return m_Email;};
  void Email(const string & value) { m_Email = value;};

  string &Contact() { return m_Contact;};
  void Contact(const string & value) { m_Contact = value;};

  string &CustomerTypeRef_ListID() { return m_CustomerTypeRef_ListID;};
  void CustomerTypeRef_ListID(const string & value) { m_CustomerTypeRef_ListID = value;};

  string &CustomerTypeRef_FullName() { return m_CustomerTypeRef_FullName;};
  void CustomerTypeRef_FullName(const string & value) { m_CustomerTypeRef_FullName = value;};

  string &TermsRef_ListID() { return m_TermsRef_ListID;};
  void TermsRef_ListID(const string & value) { m_TermsRef_ListID = value;};

  string &TermsRef_FullName() { return m_TermsRef_FullName;};
  void TermsRef_FullName(const string & value) { m_TermsRef_FullName = value;};

  string &Balance() { return m_Balance;};
  void Balance(const string & value) { m_Balance = value;};

  string &TotalBalance() { return m_TotalBalance;};
  void TotalBalance(const string & value) { m_TotalBalance = value;};

  string &SalesTaxCodeRef_ListID() { return m_SalesTaxCodeRef_ListID;};
  void SalesTaxCodeRef_ListID(const string & value) { m_SalesTaxCodeRef_ListID = value;};

  string &SalesTaxCodeRef_FullName() { return m_SalesTaxCodeRef_FullName;};
  void SalesTaxCodeRef_FullName(const string & value) { m_SalesTaxCodeRef_FullName = value;};

  string &ItemSalesTaxRef_ListID() { return m_ItemSalesTaxRef_ListID;};
  void ItemSalesTaxRef_ListID(const string & value) { m_ItemSalesTaxRef_ListID = value;};

  string &ItemSalesTaxRef_fullName() { return m_ItemSalesTaxRef_ListID;};
  void ItemSalesTaxRef_fullName(const string & value) { m_ItemSalesTaxRef_ListID = value;};

private:

  // CustomerRet Elements
  string  m_ListID;
  string  m_TimeCreated;
  string  m_TimeModified;
  string  m_EditSequence;
  string  m_Name;
  string  m_FullName;
  string  m_IsActive;

  string  m_ParentRef_ListID;
  string  m_ParentRef_FullName;

  string  m_Sublevel;
  string  m_CompanyName;
  string  m_FirstName;
  string  m_LastName;

  string  m_BillAddress_Addr1;
  string  m_BillAddress_Addr2;
  string  m_BillAddress_City;
  string  m_BillAddress_State;
  string  m_BillAddress_PostalCode;

  string  m_Phone;
  string  m_Email;
  string  m_Contact;

  string  m_CustomerTypeRef_ListID;
  string  m_CustomerTypeRef_FullName;

  string  m_TermsRef_ListID;
  string  m_TermsRef_FullName;

  string  m_Balance;
  string  m_TotalBalance;

  string  m_SalesTaxCodeRef_ListID;
  string  m_SalesTaxCodeRef_FullName;

  string  m_ItemSalesTaxRef_ListID;
  string  m_ItemSalesTaxRef_fullName;

};

#endif // !defined(AFX_QBCUSTOMERRET_H__8B2D3111_B018_45FC_80C2_C9ABEA411886__INCLUDED_)
