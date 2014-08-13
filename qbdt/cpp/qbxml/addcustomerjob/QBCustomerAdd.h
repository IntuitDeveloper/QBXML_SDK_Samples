/*---------------------------------------------------------------------------
 * FILE: QBCustomerAdd.h
 *
 * Description:
 * Customer Add operation. This class is a subclass of qbXMLOpsBase.
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

#if !defined(AFX_QBCUSTOMERADD_H__DC694611_A372_4F5A_AF6E_941E6190B068__INCLUDED_)
#define AFX_QBCUSTOMERADD_H__DC694611_A372_4F5A_AF6E_941E6190B068__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "qbXMLOpsBase.h"

#include "QBCustomerRet.h"


class QBCustomerAdd: public qbXMLOpsBase
{
public:
  QBCustomerAdd(const string &companyFileName,
                qbXMLRPWrapper *pQbxmlrpWrapper);
  virtual ~QBCustomerAdd();

  // Add customer to QuickBooks
  bool AddCustomer();

  // Set the data to send to QuickBooks
  void CustomerRet(QBCustomerRet *custRet) { m_CustomerRet = custRet;};

private:

  QBCustomerRet *m_CustomerRet;

  // helper methods
  virtual bool CreateXMLBody(); // Create CustomerAddRq Block
  virtual bool ParseXML();      // parse Response

  // No default constructor provided
  QBCustomerAdd(){};

};

#endif // !defined(AFX_QBCUSTOMERADD_H__DC694611_A372_4F5A_AF6E_941E6190B068__INCLUDED_)
