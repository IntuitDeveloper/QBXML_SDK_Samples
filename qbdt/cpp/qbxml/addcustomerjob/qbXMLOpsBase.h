/*---------------------------------------------------------------------------
 * FILE: qbXMLOpsBase.h
 *
 * Description:
 * The base class defining the basic operation of qbXMLops.
 * It provides the framework and common methods to build XML,
 * call qbXML COM, and parse response XML.
 *
 * There are 2 pure virtual methods here:
 * CreateXMLBody() -- the child class must implement this method to create
 *                    it's own request XML
 * ParseXML()      -- the child class must implement this method to parse
 *                    response XML
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

#if !defined(AFX_QBXMLOPSBASE_H__2A2010FB_6120_4616_937F_442C34EAF781__INCLUDED_)
#define AFX_QBXMLOPSBASE_H__2A2010FB_6120_4616_937F_442C34EAF781__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "DOMXMLBuilder.h"
#include "qbXMLRPWrapper.h"

class qbXMLOpsBase
{
public:

  // ctor and dtor
  qbXMLOpsBase();
  virtual ~qbXMLOpsBase();

  // status messages
  string &GetErrorMsg() { return m_ErrorMsg;};
  string &GetStatusMessage() { return m_StatusMessage;};
  string &GetStatusSeverity() { return m_StatusSeverity;};

protected:

  string   m_CompanyFileName; // QB company file name
  string   m_RequestXML;      // request XML
  string   m_ResponseXML;     // response XML

  string   m_StatusCode;      // XML status Code
  string   m_StatusSeverity;  // XML status Severity
  string   m_StatusMessage;   // XML status Message

  string   m_ErrorMsg;        // error message, set when a helper returns false

  DOMXMLBuilder  *m_Builder;  // DOMXMLBuilder
  qbXMLRPWrapper *m_pQbxmlrpWrapper;  // wrapper for qbXMLRP COM object


  // helper methods
  virtual bool BuildXML();   // create m_RequestXML
  bool CallqbXMLCOM();       // connection to qbXML COM

  // Each subclass must implement the following 2 pure virtual functions
  virtual bool ParseXML() = 0; // pure virtual function, parse Response
  virtual bool CreateXMLBody() = 0; // pure virtual function, to create
                                    // the request body

  // processes the request using BuildXML, CallqbXMLCOM, and ParseXML
  bool  PerformOperation();
};

#endif // !defined(AFX_QBXMLOPSBASE_H__2A2010FB_6120_4616_937F_442C34EAF781__INCLUDED_)
