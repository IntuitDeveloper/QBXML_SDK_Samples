/*---------------------------------------------------------------------------
 * FILE: qbXMLRPWrapper.cpp
 *
 * Description:
 * qbXMLRPWrapper Implementation, it encapsulates the details of the qbXML COM
 * calls. It provides a simple interface to call qbXML COM. The core method
 * is QBXMLRPCall, which in turn makes the qbXML COM calls
 *
 * Created On: 11/07/2001
 * Updated to QBXML 2.0 August 2002
 *
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 */
#include "stdafx.h"

#include "qbXMLWrapper.h"
#include "Utility.h"

#import "qbxmlrp.dll" named_guids //raw_interfaces_only

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

qbXMLRPWrapper::qbXMLRPWrapper(const string &companyFile,
                               const string &requestXML)
  : m_CompanyFile(companyFile),
    m_RequestXML (requestXML),
    m_ErrorStatus(false)
{

  m_Msg = "Success";

}

qbXMLRPWrapper::~qbXMLRPWrapper()
{

}


/*---------------------------------------------------------------------------
 * QBXMLRPCall
 * Here, we actually make the connection to QuickBooks, pass in our
 * request XML, get back the response XML from QuickBooks, and close
 * the connection to QuickBooks.
 *
 * We make the following COM calls to the qbXML Request Processor:
 *    1. OpenConnection
 *    2. BeginSession
 *    3. ProcessRequest
 *    4. EndSession
 *    5. CloseConnection
 *
 * Keep in mind that because the QBXMLRP is a COM object, your project
 * must call CoInitialize() somewhere before you talk to QBXMLRP,
 * and CoUninitialize() sometime before your program exits. In this example, it
 * is _tWinMain() function in AddCustomer.cpp that makes these calls.
 *
 */
string & qbXMLRPWrapper::QBXMLRPCall()
{

  bool isSessionOn = false; // in session flag
  bool isConnOpen = false; // open connection flag

  //// Initialize COM
  //if (FAILED(CoInitialize(NULL))) {
  //  m_ErrorStatus = true;
  //  m_Msg = "Unable to initialize COM";
  //}

  _bstr_t appID("SAC");
  _bstr_t appName("IDN Sample:Add Customer");
  _bstr_t bstrReqXML(m_RequestXML.c_str());
  _bstr_t bstrComFile(m_CompanyFile.c_str());
  _bstr_t ticket;
  _bstr_t bstrResXML;

  try {
    // Create RequestProcessor Ptr
    QBXMLRPLib::IRequestProcessorPtr rqPtr(QBXMLRPLib::CLSID_RequestProcessor);

    try {

      // open connection
      rqPtr ->OpenConnection(appID, appName);
      isConnOpen = true;

      // begin session
      ticket = rqPtr ->BeginSession(bstrComFile,
                                    QBXMLRPLib::qbFileOpenDoNotCare);

      isSessionOn = true;
      //process the request
      bstrResXML = rqPtr->ProcessRequest(ticket, bstrReqXML);
      Utility::BSTRToString(bstrResXML, m_ResponseXML);

      //end session and close the connection
      rqPtr->EndSession (ticket);
      isSessionOn = false;

      rqPtr->CloseConnection ();
      isConnOpen = false;

    } catch (_com_error e ){

      m_ErrorStatus = true;

      Utility::BSTRToString(e.Description(), m_Msg);

      // Make sure we end the session and close the connection
      if (isSessionOn){
        rqPtr->EndSession(ticket);
      }

      if (isConnOpen){
        rqPtr->CloseConnection ();
      }
    }
  } catch (_com_error err) {
    // Can't Instantiate IRequestProccessorPtr
    m_ErrorStatus = true;
    Utility::BSTRToString(err.Description(), m_Msg);
  }

  //CoUninitialize();

  return m_ResponseXML;

}