/*---------------------------------------------------------------------------
 * FILE: qbXMLRPWrapper.cpp
 *
 * Description:
 * qbXMLRPWrapper Implementation.  It encapsulates the details of
 * getting and releasing the IRequestProcessor COM interface,
 * and the details of the qbXMLRP COM calls.
 *
 * Only one instance of this class should be created for your entire
 * application, as each instance will ultimately open and hang on
 * to a connection to QuickBooks.  Ideally, your application should
 * only call qbXMLRP's OpenConnection()/CloseConnection() once
 * for your application's lifetime.
 *
 * Your application may want to talk to multiple QB company files.
 * When it is time to choose your first company file or switch to
 * a different one, simply call OpenCompanyFile() in this class.
 * This will call qbXMLRP's EndSession() for you if you are already
 * working with a different company file, then it will call
 * qbXMLRP's BeginSession() with the new company file you've selected,
 * and hang on to the session ticket for you.
 *
 * Once you're in a session, you may start sending requests to QB
 * by calling ProcessRequest() in this class.  This class will
 * take your request and the ticket it saved off earlier, and use
 * them to call qbXMLRP's ProcessRequest().  The response will be
 * returned to you and it is your job to parse it.
 *
 * The destructor for this class will call qbXMLRP's EndSession() and
 * CloseConnection() for you, if needed, releasing your connection
 * to QuickBooks.  Thus, you should only call the destructor if
 * your application no longer needs to communicate to QB for the
 * remainder of its lifetime.
 *
 * In summary, we encapsulate the following qbXMLRP calls:
 *    OpenConnection()
 *    BeginSession()
 *    ProcessRequest()
 *    ProcessSubscriptionRequest()
 *    EndSession()
 *    CloseConnection()
 *
 * Created On: 11/11/2001
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 */
#include "stdafx.h"
#include "qbXMLRPWrapper.h"

#include "Utility.h"

#import "qbxmlrp2.dll" no_namespace named_guids raw_interfaces_only

static const char *THE_APPID = "";
static const char *THE_APPNAME = "QBSDK Event and UI Sample";


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////


/*---------------------------------------------------------------------------
 * Constructor
 *
 */
qbXMLRPWrapper::qbXMLRPWrapper()
  : m_HasError(false),
    m_IsConnectionOpen(false),
    m_pReqProc(NULL)
{
  m_ErrorMsg = "OK";
}


/*---------------------------------------------------------------------------
 * Destructor
 *
 */
qbXMLRPWrapper::~qbXMLRPWrapper()
{
  // Call EndSession if BeginSession was called.
  CloseCompanyFile();

  // Call CloseConnection if OpenConnection was called.
  CloseQBConnection();

  // Since we have a raw COM pointer, it is important to Release() it.
  if (m_pReqProc != NULL) {
    m_pReqProc->Release();
    m_pReqProc = NULL;
  }
}


/*---------------------------------------------------------------------------
 * OpenCompanyFile
 * Call this to connect to a QuickBooks company file.
 *
 * If you are already connected to the same file, that's fine...
 * we just won't do anything.
 *
 * If you are already connected to a different file, we will end that
 * session first, then begin a new session with the new file.
 *
 * If you're not connected to QB yet, then we will open the connection
 * to QB and begin a new session with the specified company file.
 *
 */
bool qbXMLRPWrapper::OpenCompanyFile(string companyFile)
{
  HRESULT hr = S_OK;;
  CComBSTR ticket;

  // Reset error state
  m_HasError = false;
  m_ErrorMsg = "OK";

  // Get an instance of the qbXMLRP Request Processor and
  // call OpenConnection if that has not been done already.
  if (!OpenQBConnection()) {
    return false;
  }

  // If the requested company file is already open, we need not do anything.
  if ( companyFile.length() == 0 || companyFile != m_CompanyFile) {
    // If we are in another session, call EndSession first.
    if( companyFile.length() > 0 ) {
        if (!CloseCompanyFile()) {
          return false;
        }
    }

    // Now we can call BeginSession for this company file.
	if( m_Ticket.empty() ) {
	    hr = m_pReqProc ->BeginSession(_bstr_t(companyFile.c_str()),
		                               qbFileOpenDoNotCare,
			                           &ticket);
        if (FAILED (hr)) {
          HandleCOMError();
          return false;
        }
        else {
            // Remember the company we're working for and the session ticket.
            m_CompanyFile = companyFile;
            Utility::BSTRToString(ticket, m_Ticket);
        }
	}

  }

  return true;
}


/*---------------------------------------------------------------------------
 * CloseCompanyFile
 * Private helper method that will call EndSession if we're currently
 * in a session.
 *
 */
bool qbXMLRPWrapper::CloseCompanyFile() {
  HRESULT hr = S_OK;

  if (!m_Ticket.empty() && !m_CompanyFile.empty()) {
    hr = m_pReqProc->EndSession (_bstr_t(m_Ticket.c_str()));
    if (FAILED (hr)) {
      HandleCOMError();
      return false;
    }

    m_Ticket.erase();
    m_CompanyFile.erase();
  }

  return true;
}


/*---------------------------------------------------------------------------
 * OpenQBConnection
 * Private helper method that will get an instance of the qbXML
 * Request Processor, if needed, and call OpenConnection if no
 * connection has been established yet.
 *
 */
bool qbXMLRPWrapper::OpenQBConnection() {
  HRESULT hr = S_OK;

  if (!m_IsConnectionOpen) {
    if (m_pReqProc == NULL) {
      // We need an instance of the qbXML Request Processor.
      // This is our communication channel to QuickBooks.
      HRESULT hr;
      hr = CoCreateInstance(CLSID_RequestProcessor2,
                            NULL,
                            CLSCTX_INPROC_SERVER,
                            IID_IRequestProcessor3,
                            reinterpret_cast<void**>(&m_pReqProc));
      if (FAILED(hr) || (m_pReqProc == NULL)) {
        TCHAR tmpMsg[1024];
        sprintf (tmpMsg, "Failed to instantiate the qbXML Request Processor.  Error: 0x%lx.  Perhaps qbxmlrp2.dll is missing?", hr);
        m_ErrorMsg = tmpMsg;
        m_HasError = true;
        return false;
      }
    }

    // Please see the documentation for a discussion of what to choose
    // for the appID and appName parameters.
    hr = m_pReqProc ->OpenConnection(_bstr_t(THE_APPID), _bstr_t(THE_APPNAME));
    if (FAILED (hr)) {
      HandleCOMError();
      return false;
    }

    m_IsConnectionOpen = true;
  }

  return true;
}


/*---------------------------------------------------------------------------
 * CloseQBConnection
 * Private helper method that will call CloseConnection if a
 * connection to QB is currently open.
 *
 */
bool qbXMLRPWrapper::CloseQBConnection() {
  HRESULT hr = S_OK;

  if (m_IsConnectionOpen) {
    hr = m_pReqProc->CloseConnection ();
    if (FAILED (hr)) {
      HandleCOMError();
      return false;
    }

    m_IsConnectionOpen = false;
  }

  return true;
}


/*---------------------------------------------------------------------------
 * HandleCOMError
 * Private helper method that will get the IErrorInfo from the
 * last COM error generated by qbXMLRP, and set both the m_ErrorMsg
 * and m_HasError data members.
 *
 */
void qbXMLRPWrapper::HandleCOMError() {
  IErrorInfo    *pError;
  HRESULT       errorHr = S_OK;

  m_HasError = true;

  errorHr = GetErrorInfo (0, &pError);

  // Note that we explicitly check for S_OK because GetErrorInfo
  // can return S_FALSE when there's no error object to return,
  // which would pass the SUCCEEDED() test.
  if (errorHr == S_OK  &&  pError != NULL) {
    CComBSTR    description;

    try {
      pError->GetDescription (&description);

      Utility::BSTRToString(description, m_ErrorMsg);
    }
    catch (const _com_error& ex) {
      ex.Error();
      m_ErrorMsg = "Error getting error description.";
    }
  }
  else {
    m_ErrorMsg = "Error getting error info.";
  }
}


/*---------------------------------------------------------------------------
 * ProcessRequest
 * Call this to send a qbXML request to Quickbooks to the currently
 * opened QB company file.  We assume you have already called
 * OpenCompanyFile().
 *
 * A qbXML response will be returned unless an internal error occurred.
 * You should check the GetHasError() method for any internal errors,
 * and if so, call GetErrorMsg() for the corresponding error message.
 * If there are no internal errors, then you should parse the qbXML
 * response to get the status of your processed request.
 *
 */
string &qbXMLRPWrapper::ProcessRequest(string requestXML)
{
  HRESULT hr = S_OK;

  // Reset error state
  m_HasError = false;
  m_ErrorMsg = "OK";

  CComBSTR bstrTicket(m_Ticket.c_str());
  CComBSTR bstrReqXML(requestXML.c_str());
  CComBSTR bstrComFile(m_CompanyFile.c_str());

  BSTR bstrResXML = NULL;

  hr = m_pReqProc->ProcessRequest(bstrTicket.m_str,
                                  bstrReqXML.m_str,
                                  &bstrResXML);
  if (FAILED (hr)) {
    HandleCOMError();
  }

  Utility::BSTRToString(bstrResXML, m_ResponseXML);

  return m_ResponseXML;
}

/*---------------------------------------------------------------------------
 * ProcessSubscription
 * Call this to send a qbXML subscription request to Quickbooks.
 *
 * A qbXML response will be returned unless an internal error occurred.
 * You should check the GetHasError() method for any internal errors,
 * and if so, call GetErrorMsg() for the corresponding error message.
 * If there are no internal errors, then you should parse the qbXML
 * response to get the status of your processed request.
 *
 */
string &qbXMLRPWrapper::ProcessSubscription(string requestXML)
{
  HRESULT hr = S_OK;

  // Reset error state
  m_HasError = false;
  m_ErrorMsg = "OK";

  // Get an instance of the qbXMLRP Request Processor and
  // call OpenConnection if that has not been done already.
  if (!OpenQBConnection()) {
	m_ResponseXML = "";
	return m_ResponseXML;
  }

  CComBSTR bstrReqXML(requestXML.c_str());

  BSTR bstrResXML = NULL;

  hr = m_pReqProc->ProcessSubscription( bstrReqXML.m_str, &bstrResXML );
  if (FAILED (hr)) {
    HandleCOMError();
  }

  Utility::BSTRToString(bstrResXML, m_ResponseXML);

  return m_ResponseXML;
}