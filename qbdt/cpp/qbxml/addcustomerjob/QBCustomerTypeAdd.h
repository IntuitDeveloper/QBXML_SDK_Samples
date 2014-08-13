/*---------------------------------------------------------------------------
 * FILE: QBCustomerTypeAdd.h
 *
 * Description:
 * CustomerType Add operation. This class is a subclass of qbXMLOpsBase.
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

#if !defined(AFX_QBCustomerTypeTYPEADD_H__43F6B047_A373_4356_9ED5_40573BFA57C3__INCLUDED_)
#define AFX_QBCustomerTypeTYPEADD_H__43F6B047_A373_4356_9ED5_40573BFA57C3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "qbXMLOpsBase.h"

#include "QBCustomerTypeRet.h"


class QBCustomerTypeAdd : public qbXMLOpsBase
{
public:
  QBCustomerTypeAdd(const string &companyFileName,
                    qbXMLRPWrapper *pQbxmlrpWrapper);
  virtual ~QBCustomerTypeAdd();

  // Add customer type to QuickBooks
  bool AddCustomerType();

  // Set the data to send to QuickBooks
  void CustomerTypeRet(QBCustomerTypeRet *custRet) { m_CustomerTypeRet = custRet;};

private:

  QBCustomerTypeRet *m_CustomerTypeRet;

  // helper methods
  virtual bool CreateXMLBody(); // Create CustomerTypeAddRq Block
  virtual bool ParseXML();      // parse Response

  // No default constructor provided
  QBCustomerTypeAdd(){};

};

#endif // !defined(AFX_QBCustomerTypeTYPEADD_H__43F6B047_A373_4356_9ED5_40573BFA57C3__INCLUDED_)
