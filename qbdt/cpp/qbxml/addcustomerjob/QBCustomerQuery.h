/*---------------------------------------------------------------------------
 * FILE: QBCustomerQuery.h
 *
 * Description:
 * Customer Query operation. This class is a subclass of qbXMLOpsBase.
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

#if !defined(AFX_QBCUSTOMERQUERY_H__E8F443ED_7365_45A9_ABE0_A5E7E1285D4C__INCLUDED_)
#define AFX_QBCUSTOMERQUERY_H__E8F443ED_7365_45A9_ABE0_A5E7E1285D4C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "qbXMLOpsBase.h" // base class

#include "QBCustomerRet.h"


typedef vector<QBCustomerRet*> QBCustRetVec;

class QBCustomerQuery : public qbXMLOpsBase
{
public:
  QBCustomerQuery(const string &companyFileName,
                  qbXMLRPWrapper *pQbxmlrpWrapper);
  virtual ~QBCustomerQuery();

  // main retrieval method
  bool GetCustomerListFromQB();

  // access the the CustomerRet Array
  QBCustRetVec &GetCustomerVector() { return m_CustRetVec;};

  // Customer Count
  int Count() { return m_CustomerCount;};

private:

  int m_CustomerCount;

  // vector m_CustRetVec is used to store customer list
  QBCustRetVec m_CustRetVec;

  // helper methods
  virtual bool CreateXMLBody(); // Create CustomerQueryRq Block
  virtual bool ParseXML();      // parse Response

  // No default constructor provided
  QBCustomerQuery(){};

};

#endif // !defined(AFX_QBCUSTOMERQUERY_H__E8F443ED_7365_45A9_ABE0_A5E7E1285D4C__INCLUDED_)
