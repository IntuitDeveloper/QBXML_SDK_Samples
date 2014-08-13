/*---------------------------------------------------------------------------
 * FILE: QBCustomerTypeQuery.h
 *
 * Description:
 * CustomerType Query operation. This class is a subclass of qbXMLOpsBase.
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

#if !defined(AFX_QBCUSTOMERTYPEQUERY_H__7FA0C3D2_F70C_4584_BFBC_BBC83CBB6E4F__INCLUDED_)
#define AFX_QBCUSTOMERTYPEQUERY_H__7FA0C3D2_F70C_4584_BFBC_BBC83CBB6E4F__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "qbXMLOpsBase.h"

#include "QBCustomerTypeRet.h"


typedef vector<QBCustomerTypeRet*> QBCustTypeRetVec; // CustomerType Vector

class QBCustomerTypeQuery : public qbXMLOpsBase
{
public:
  QBCustomerTypeQuery(const string &companyFileName,
                      qbXMLRPWrapper *pQbxmlrpWrapper);
  virtual ~QBCustomerTypeQuery();

  // main retrieval method
  bool GetCustomerTypeListFromQB();

  // access the the CustomerRet Vector
  QBCustTypeRetVec &GetCustomerTypeVector() { return m_CustTypeRetVec;};

  // CustomerType Count
  int Count() { return m_CustomerTypeCount;};

private:

  int m_CustomerTypeCount;

  // m_CustTypeRetVec is used to store customer type list
  QBCustTypeRetVec m_CustTypeRetVec;

  // helper methods
  virtual bool CreateXMLBody(); // Create CustomerTypeQueryRq Block
  virtual bool ParseXML();      // parse Response

  // No default constructor provided
  QBCustomerTypeQuery(){};

};

#endif // !defined(AFX_QBCUSTOMERTYPEQUERY_H__7FA0C3D2_F70C_4584_BFBC_BBC83CBB6E4F__INCLUDED_)
