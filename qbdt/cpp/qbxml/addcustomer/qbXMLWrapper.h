/*---------------------------------------------------------------------------
 * FILE: qbXMLRPWrapper.h
 *
 * Description:
 * qbXMLRPWrapper Header file
 *
 * Created On: 11/07/2001
 * Updated to QBXML 2.0 August 2002
 *
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 */

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class qbXMLRPWrapper
{
public:

  // Constructor
  qbXMLRPWrapper(const string &companyFile,
                 const string &requestXML);

  // Destructor
  virtual ~qbXMLRPWrapper();

  // Get methods
  string & GetRequestXML() { return m_RequestXML;};
  string & GetMsg() { return m_Msg;};

  bool     GetErrorStatus() { return m_ErrorStatus;};

  // Core qbXML COM Connection
  //
  string & QBXMLRPCall();

private:

  string m_RequestXML;  // request XML
  string m_ResponseXML; // response XML
  string m_CompanyFile; // company file

  string m_Msg; // message
  bool   m_ErrorStatus;  // error Status, if true means there is an error

  // no default constructor is provided
  qbXMLRPWrapper(){};
};

