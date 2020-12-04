/*---------------------------------------------------------------------------
 * FILE: AddCustDlg.cpp
 *
 * Description:
 * ATL Dlg Class AddCustDlg Implementation
 *
 * Created On: 11/07/2001
 * Updated to QBXML 2.0 August 2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */
#include "stdafx.h"
#include "AddCustDlg.h"

#include "ValidatorDlg.h"
#include "DisplayDlg.h"

// Constructor
//
CAddCustDlg::CAddCustDlg()
{

  m_CompanyFile[0] = '\0';
  m_CustName[0]    = '\0';
  m_Phone[0]       = '\0';
  m_FirstName[0]   = '\0';
  m_LastName[0]    = '\0';

}

// Destructor
//
CAddCustDlg::~CAddCustDlg()
{

}

/*---------------------------------------------------------------------------
 * OnClickedSubmit
 * Add Customer button event handler, this is the main method of the application.
 * It performs the following steps:
 *     1. Collect form data
 *     2. Build request XML
 *     3. Call qbXMLRP
 *     4. Parse the response XML
 *
 */
LRESULT CAddCustDlg::OnClickedSubmit(WORD   wNotifyCode,
                                     WORD   wID,
                                     HWND   hWndCtl,
                                     BOOL   & bHandled)
{

  // Collect data from the Dlg.
  //
  if (!CollectFormData(_T("'Add Customer'."))) {
    MessageBox(m_Msg.c_str(), _T("Data Missing"));
    return 0;
  }

  // create a CustomerAdd object
  //
  CCustomerAdd cAdd = CCustomerAdd(m_CompanyFile,
                                   m_CustName,
                                   m_FirstName,
                                   m_LastName,
                                   m_Phone);

  // build XML
  cAdd.BuildXML();
  if (cAdd.GetErrorStatus()) { // build failed
    MessageBox(cAdd.GetMsg().c_str(), _T("Build XML Failed"));
    return 0;
  }

  // set responseXML
  //
  m_RequestXML = cAdd.GetRequestXML();

  // Call qbXMLRP
  //
  cAdd.DoRequest();
  if (cAdd.GetErrorStatus()) { // call qbXML failed
    MessageBox(cAdd.GetMsg().c_str(), "Call QBXMLRP Failed");
    return 0;
  }

  // set response XML
  //
  m_ResponseXML = cAdd.GetResponseXML();

  //Parse Response XML
  //
  cAdd.ParseResponseXML();
  if (cAdd.GetErrorStatus()) { // call qbXML failed
    MessageBox(cAdd.GetMsg().c_str(), "Parse XML Failed");
    return 0;
  }

  // display the status message from the response XML
  if (strncmp(cAdd.GetStatusSeverity().c_str(), "Error", 5) != 0) {
    MessageBox(cAdd.GetStatusMessage().c_str(), "Add Customer Status");
  }
  else {
    MessageBox(cAdd.GetStatusMessage().c_str(), "Customer Add Error");
  }

  return 0;
}

/*---------------------------------------------------------------------------
 * OnClickedView_req
 * View XML Request button event handler. Build and display the request XML
 * based on the form data that the user enterred.
 *
 */
LRESULT CAddCustDlg::OnClickedView_req(WORD   wNotifyCode,
                                       WORD   wID,
                                       HWND   hWndCtl,
                                       BOOL   & bHandled)
{
  // Show Request XML
  //

  // Collect data
  if (!CollectFormData(_T("'View XML Request'."))) {
    MessageBox(m_Msg.c_str(), _T("Data Missing"));
    return 0;
  }

  // Create a new CustomerAdd Instance
  //
  CCustomerAdd custAdd = CCustomerAdd(m_CompanyFile,
                                      m_CustName,
                                      m_FirstName,
                                      m_LastName,
                                      m_Phone);

  custAdd.BuildXML();
  m_RequestXML = custAdd.GetRequestXML();

  // Create the Display dialog window
  //
  CDisplayDlg cDisplayXMLDlg(m_RequestXML);


  HWND pWin = ::GetActiveWindow();
  int ret = cDisplayXMLDlg.DoModal(pWin);

  MSG msg;

  // to make tab work
  while (GetMessage(&msg, 0, 0, 0)) {
    if ((cDisplayXMLDlg) &&
       (!::IsDialogMessage(cDisplayXMLDlg.m_hWnd,&msg))) {
      DispatchMessage(&msg);
    }
  }
  return 0;
}


/*---------------------------------------------------------------------------
 * OnClickedView_res
 * View QB Response button event handler. Show the response XML
 *
 */
LRESULT CAddCustDlg::OnClickedView_res(WORD   wNotifyCode,
                                       WORD   wID,
                                       HWND   hWndCtl,
                                       BOOL   & bHandled)
{
  // Show Response XML
  if (!m_ResponseXML.empty()) {

    // replace "\n" with "\n\r"
    string newResponseXML;
    ReplaceChar(m_ResponseXML, newResponseXML);

    // Create the Display dialog window
    //
    CDisplayDlg cDisplayXMLDlg(newResponseXML);

    HWND pWin = ::GetActiveWindow();
    int ret = cDisplayXMLDlg.DoModal(pWin);

    MSG msg;

    // to make tab work
    while (GetMessage(&msg, 0, 0, 0)) {
      if ((cDisplayXMLDlg) &&
         (!::IsDialogMessage(cDisplayXMLDlg.m_hWnd,&msg))) {
        DispatchMessage(&msg);
      }
    }
  }
  else {
    MessageBox(_T("Please make a request before checking the response."), _T("Response XML"));
  }
  return 0;
}

/*---------------------------------------------------------------------------
 * OnClickedValidateReq
 * Check XML Request button event handler, Show the XML validation result
 *
 */
LRESULT CAddCustDlg::OnClickedValidateReq(WORD  wNotifyCode,
                                          WORD  wID,
                                          HWND  hWndCtl,
                                          BOOL  & bHandled)
{

  // Get Request XML
  //
  // Collect data
  if (!CollectFormData(_T("'Check XML Request'."))) {
    MessageBox(m_Msg.c_str(), _T("Data Missing"));
    return 0;
  }

  // Create a new CustomerAdd Instance
  //
  CCustomerAdd custAdd = CCustomerAdd(m_CompanyFile,
                                      m_CustName,
                                      m_FirstName,
                                      m_LastName,
                                      m_Phone);
  custAdd.BuildXML();
  m_RequestXML = custAdd.GetRequestXML();

  // Create the Validator dialog window
  //
  CValidatorDlg cValXMLDlg(m_RequestXML);


  HWND pWin = ::GetActiveWindow();
  int ret = cValXMLDlg.DoModal(pWin);

  cValXMLDlg.SetModalReturnCode(ret);

  MSG msg;

  // to make tab work
  while (GetMessage(&msg, 0, 0, 0)) {
    if ((cValXMLDlg) &&
       (!::IsDialogMessage(cValXMLDlg.m_hWnd,&msg))) {
      DispatchMessage(&msg);
    }
  }

  return 0;
}

/*---------------------------------------------------------------------------
 * CollectFormData
 * Collect form data.  Returns false if a required field on the form is empty.
 *
 */
bool CAddCustDlg::CollectFormData(TCHAR * szRestOfMsg)
{

  // Company File Name
  this -> GetDlgItemText(IDC_COM_FILE, (LPTSTR) m_CompanyFile, FIELD_SIZE - 1);

  // Customer Name, is required
  this -> GetDlgItemText(IDC_CUST_NAME, (LPTSTR) m_CustName, FIELD_SIZE - 1);
  if (m_CustName[0] == '\0') {
    m_Msg = "Please enter a customer name to add and then click ";
    m_Msg += szRestOfMsg;
    return false;
  }
  // Phone, is optional
  this -> GetDlgItemText(IDC_PHONE, (LPTSTR) m_Phone, FIELD_SIZE - 1);

  // First Name, is optional
  this -> GetDlgItemText(IDC_FIRSTNAME, (LPTSTR) m_FirstName, FIELD_SIZE - 1);

  // Last Name, is optional
  this -> GetDlgItemText(IDC_LASTNAME, (LPTSTR) m_LastName, FIELD_SIZE - 1);

  return true;
}

/*---------------------------------------------------------------------------
 * OnClickedBrowse
 * Brings up a standard File Open dialog
 *
 */
LRESULT CAddCustDlg::OnClickedBrowse(WORD   wNotifyCode,
                                     WORD   wID,
                                     HWND   hWndCtl,
                                     BOOL   & bHandled)
{
  OPENFILENAME of;
  TCHAR        szFile[MAX_PATH];

  GetDlgItemText(IDC_COM_FILE, szFile, MAX_PATH);
  memset(&of, 0, sizeof(OPENFILENAME));

  of.lStructSize  = sizeof(OPENFILENAME);
  of.hwndOwner    = this->m_hWnd;
  of.lpstrFilter  = _T("QuickBooks Data Files (*.QBW)\0*.QBW\0\0"); ;
  of.nFilterIndex = 0L;
  of.lpstrFile    = szFile;
  of.nMaxFile     = MAX_PATH;
  of.lpstrTitle   = "Browse for QuickBooks File";
  of.Flags        =  OFN_SHOWHELP       |
                     OFN_PATHMUSTEXIST  |
                     OFN_FILEMUSTEXIST  |
                     OFN_HIDEREADONLY   |
                     OFN_EXPLORER       |
                     OFN_LONGNAMES;

  if (GetOpenFileName(&of)) {
    SetDlgItemText(IDC_COM_FILE, of.lpstrFile);
  }

  return 0L;
}

/*---------------------------------------------------------------------------
 * InitFileName
 * Sets up path to company file based on the registry.
 *
 */
TCHAR * CAddCustDlg::InitFileName(TCHAR * szDefaultPath,
                                  TCHAR * szFileName)
{
  HKEY    hKey = NULL;
  static  TCHAR szPath[FIELD_SIZE];  // make it static so we can return it

  szPath[0] = '\0';

  return szPath;
}

/*---------------------------------------------------------------------------
 * ReplaceChar
 * Utility function to replace "\n" to "\n\r" just for window
 * display purpose.
 *
 */
void CAddCustDlg::ReplaceChar(const string  &inStr,
                              string        &outStr)
{
  if (inStr.empty()){
    return;
  }

  const TCHAR   newline = 0x0a; // nl
  const string  newlineCReturn = "\r\n";

  int len = inStr.size();

  TCHAR cChar;

  for (int i = 0; i < len; ++i) {
    cChar = inStr[i];

    if (cChar == newline) { // replace the char
      outStr += newlineCReturn;
    }
    else {
      outStr += cChar;
    }
  }

}
