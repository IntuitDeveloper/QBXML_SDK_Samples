/*---------------------------------------------------------------------------
 * FILE: CustomerAdd.h
 *
 * Description:
 * Header file of CustomerAdd class.
 * This class generates the XML for CustomerAdd request via XMLBuilder.
 * It also sends the qbXML request to QuickBooks via qbXMLRP COM object.
 * Then it can parse the response from QuickBooks.
 * Accessors are provided for both the request and response XML,
 * as well as individual fields of the parsed response.
 *
 * Created On: 11/07/2001
 * Updated to QBXML 2.0 August 2002
 *
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 */
#if !defined(AFX_CUSTOMERADD_H__7488CE63_CB4F_49D0_8BA9_667DD7811104__INCLUDED_)
#define AFX_CUSTOMERADD_H__7488CE63_CB4F_49D0_8BA9_667DD7811104__INCLUDED_


#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "XMLBuilder.h"

// MSXML lib
#include "msxml.h"
#import "msxml4.dll" named_guids raw_interfaces_only

class CCustomerAdd
{
public:

  // ctor and dtor
  //

  CCustomerAdd(const string &companyFileName,
               const string &custName,
               const string &firstName,
               const string &lastName,
               const string &phone);


  virtual ~CCustomerAdd();

  // 1. Build XML request
  //
  void BuildXML();

  // 2. Send the request to qbXML COM
  //
  void DoRequest();

  // 3. Parse XML response
  //
  void ParseResponseXML();

  // data access methods
  //
  string & GetRequestXML() { return m_RequestXML;};
  string & GetResponseXML() { return m_ResponseXML; };
  string & GetMsg() {return m_Msg;};
  string & GetStatusMessage() { return m_StatusMessage;};
  string & GetStatusSeverity() { return m_StatusSeverity;};

  bool     GetErrorStatus() {return m_ErrorStatus;};

private:

  // Members
  //

  XMLBuilder *m_Builder;    // XMLBuilder Ptr

  string m_CompanyFileName; // company file name
  string m_CustomerName;    // customer name
  string m_Phone;           // phone
  string m_FirstName;       // first name
  string m_LastName;        // last name

  string m_RequestXML;      // request XML
  string m_ResponseXML;     // response XML

  string m_StatusCode;      // XML status Code
  string m_StatusSeverity;  // XML status Severity
  string m_StatusMessage;   // XML status Message

  string m_Msg;             // message
  bool   m_ErrorStatus;     // error status, false means No Error

  // helper methods
  //
  void CreateCustomerAddRq(); // Create CustomerAddRq Block
  void CreateXMLTrailer();    // Create XML Trailer
  void CreateXMLHeader();     // Create XML Header

  // No default constructor provided
  CCustomerAdd(){};
};


#endif // !defined(AFX_CUSTOMERADD_H__7488CE63_CB4F_49D0_8BA9_667DD7811104__INCLUDED_)
