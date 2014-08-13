/*---------------------------------------------------------------------------
 * FILE: qbXMLOpsBase.cpp
 *
 * Description:
 * Implementation of qbXMLOpsBase
 *
 * Created On: 11/11/2001
 * Updated to SDK 2.0: 08/08/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#include "stdafx.h"

#include "qbXMLOpsBase.h"
#include "qbXMLRPWrapper.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

qbXMLOpsBase::qbXMLOpsBase()
{
  m_StatusMessage = "OK!";
  m_ErrorMsg      = "OK!";

  m_Builder = new DOMXMLBuilder();
}

qbXMLOpsBase::~qbXMLOpsBase()
{
  delete m_Builder;
}


/*--------------------------------------------------------------------
 * PerformOperation
 * Each subclass can call this method to perform the common operations
 * of building the request XML, communicating with the qbXML request
 * processor, and parsing the response XML.
 *
 */
bool qbXMLOpsBase::PerformOperation()
{
  // Build request XML
  if (!BuildXML()) {
    return false;
  }

  // Send the request XML to the qbXML request processor and get
  // the response XML back
  if (!CallqbXMLCOM()) {
    return false;
  }

  // Parse the response XML
  if (!ParseXML()) {
    return false;
  }

  return true;

}


/*--------------------------------------------------------------------
 * BuildXML
 * Build the XML request.  It involves 3 steps:
 * 1. Create Header, this is the same for all qbXML operations
 * 2. CreateXMLBody, each qbXML operation (implemented by each subclass)
 *          has its own unique request message set format
 * 3. Create Trailer, this is the same for all qbXML operations
 *
 */
bool qbXMLOpsBase::BuildXML()
{
  // Create QBXML Header
  m_Builder -> CreateXMLHeader();
  if (m_Builder ->GetHasError()) {
    m_ErrorMsg = m_Builder ->GetErrorMsg();
    return false;
  }

  // Create QBXML Body Block
  if (!CreateXMLBody()) {
    return false;
  }

  // Create QBXML trailer
  m_Builder -> CreateXMLTrailer();
  if (m_Builder ->GetHasError()) {
    m_ErrorMsg = m_Builder ->GetErrorMsg();
    return false;
  }

  // get the XML
  m_RequestXML = m_Builder ->GetXML();
  if (m_RequestXML == "") {
    m_ErrorMsg = "XML Builder failed to generate request XML!";
    return false;
  }

  return true;
}


/*--------------------------------------------------------------------
 * CallqbXMLCOM
 * We use the qbXMLRPWrapper to connect to qbXML COM.  It processes
 * our request XML and returns the response XML.
 *
 */
bool qbXMLOpsBase::CallqbXMLCOM()
{
  if (!m_pQbxmlrpWrapper ->GetHasError()) {
    m_ResponseXML = m_pQbxmlrpWrapper ->ProcessRequest(m_RequestXML);
  }

  bool hasError;
  hasError = m_pQbxmlrpWrapper ->GetHasError();
  m_ErrorMsg = m_pQbxmlrpWrapper ->GetErrorMsg();

  return !hasError;
}
