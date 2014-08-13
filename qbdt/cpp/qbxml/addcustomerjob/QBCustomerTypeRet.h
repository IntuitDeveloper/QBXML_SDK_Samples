/*---------------------------------------------------------------------------
 * FILE: QBCustomerTypeRet.h
 *
 * Description:
 * Provides get/set methods for CustomerTypeRet object
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

#if !defined(AFX_QBCUSTOMERTYPERET_H__77C8CD11_5DFB_4B41_81F1_EEF75CC2BB67__INCLUDED_)
#define AFX_QBCUSTOMERTYPERET_H__77C8CD11_5DFB_4B41_81F1_EEF75CC2BB67__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class QBCustomerTypeRet
{
public:
  QBCustomerTypeRet();
  ~QBCustomerTypeRet();

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

  string &Sublevel() { return m_Sublevel;};
  void Sublevel(const string & value) { m_Sublevel = value;};

private:
  // CustomerTypeRet Elements
  string  m_ListID;
  string  m_TimeCreated;
  string  m_TimeModified;
  string  m_EditSequence;
  string  m_Name;
  string  m_FullName;
  string  m_IsActive;
  string  m_Sublevel;
};

#endif // !defined(AFX_QBCUSTOMERTYPERET_H__77C8CD11_5DFB_4B41_81F1_EEF75CC2BB67__INCLUDED_)
