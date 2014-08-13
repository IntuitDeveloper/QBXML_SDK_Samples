/*---------------------------------------------------------------------------
 * FILE: qbXMLRPWrapper.h
 *
 * Description:
 * qbXMLRPWrapper Header file
 *
 * Created On: 11/11/2001
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 */
#if !defined(AFX_qbXMLRPWRAPPER_H__CD3AA532_5948_4A7B_ACD4_EE2FF6C4AC4D__INCLUDED_)
#define AFX_qbXMLRPWRAPPER_H__CD3AA532_5948_4A7B_ACD4_EE2FF6C4AC4D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// forward declaration
interface IRequestProcessor3;

class qbXMLRPWrapper
{
public:

  // Constructor
  qbXMLRPWrapper();

  // Destructor
  ~qbXMLRPWrapper();

  // Connect to QB and begin a session with given QB company file
  bool    OpenCompanyFile(string companyFile);

  // Process given qbXML request with currently open company file
  string  &ProcessRequest(string requestXML);

  // Process given qbXML for a subscription request.
  string  &ProcessSubscription(string requestXML);

  // Getters for error state and error message
  bool    GetHasError()  { return m_HasError; }
  string  &GetErrorMsg() { return m_ErrorMsg; }

private:

  string m_ResponseXML; // response XML
  string m_CompanyFile; // company file
  string m_Ticket;      // ticket for a QB session

  bool   m_IsConnectionOpen;  // Do we have an open connection to QB?

  bool   m_HasError;    // Is there an error?
  string m_ErrorMsg;    // error message

  IRequestProcessor3 *m_pReqProc;  // the qbXML Request Processor

  bool CloseCompanyFile();
  bool OpenQBConnection();
  bool CloseQBConnection();
  void HandleCOMError();
};

#endif // !defined(AFX_qbXMLRPWRAPPER_H__CD3AA532_5948_4A7B_ACD4_EE2FF6C4AC4D__INCLUDED_)
