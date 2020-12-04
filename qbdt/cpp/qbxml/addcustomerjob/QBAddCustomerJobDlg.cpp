/*---------------------------------------------------------------------------
 * FILE: QBAddCustomerJobDlg.cpp
 *
 * Description:
 * ATL Dlg Class CQBAddCustomerJobDlg Implementation
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
#include "QBAddCustomerJobDlg.h"
#include "QBCustomerQuery.h"
#include "QBCustomerTypeQuery.h"
#include "QBCustomerAdd.h"
#include "QBCustomerTypeAdd.h"
#include "Utility.h"

/////////////////////////////////////////////////////////////////////////////
// CQBAddCustomerJobDlg

/*---------------------------------------------------------------------------
 * Constructor
 *
 */
CQBAddCustomerJobDlg::CQBAddCustomerJobDlg(TCHAR *companyFile, qbXMLRPWrapper *pQbxmlrpWrapper)
  :m_CompanyFileName(companyFile),
   m_pQbxmlrpWrapper(pQbxmlrpWrapper),
   m_NewCustomer(NULL)
{
}


/*---------------------------------------------------------------------------
 * Destructor
 *
 */
CQBAddCustomerJobDlg::~CQBAddCustomerJobDlg()
{
  if (m_NewCustomer != NULL) {
    delete m_NewCustomer;
  }
}


/*---------------------------------------------------------------------------
 * OnInitDialog
 * This method sets up the combo box values.
 *
 * It performs the following steps:
 *     1. Get the customer list from QuickBooks by issuing a CustomerQueryRq.
 *     2. Use the results from the CustomerQueryRs to populate the
 *        parent customer combo box.
 *     3. Get the customer type list from QuickBooks by issuing a
 *        CustomerTypeQueryRq.
 *     4. Use the results from the CustomerTypeQueryRs to populate the
 *        customer type combo box.
 *
 */
LRESULT CQBAddCustomerJobDlg::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
  // ----------------------------------------------------------------------------
  // Step 1: Get the customer list from QuickBooks
  //
  QBCustomerQuery custQuery(m_CompanyFileName.c_str(),
                            m_pQbxmlrpWrapper);

  if (!custQuery.GetCustomerListFromQB()) {
    MessageBox(custQuery.GetErrorMsg().c_str(), "Error getting customer list from QuickBooks");

    // close the window
    EndDialog(m_NRetCode);
    PostQuitMessage(0);
    return 0;
  }

  // Status check
  if (custQuery.GetStatusSeverity().substr(0, 5) == "Error") {
    MessageBox(custQuery.GetStatusMessage().c_str(), "Received error when getting customer list from QuickBooks");

    // close the window
    EndDialog(m_NRetCode);
    PostQuitMessage(0);
    return 0;
  }

  // ----------------------------------------------------------------------------
  // Step 2: Populate the parent customer combo box.
  //

  QBCustRetVec custRetVec = custQuery.GetCustomerVector();
  QBCustRetVec::iterator it;

  bool gotCurSel = false;
  for (it = custRetVec.begin(); it != custRetVec.end(); it++){

    QBCustomerRet *custRet = *it;

    // Add a customer to the combo box
    SendDlgItemMessage(IDC_COMBO_CustList, CB_ADDSTRING, 0, (LPARAM)_T(custRet ->FullName().c_str()));

    // Set the first customer as the currently selected value
    if (!gotCurSel ) {
      SendDlgItemMessage(IDC_COMBO_CustList, CB_SETCURSEL, 0, (LPARAM)_T(custRet ->FullName().c_str()));
      gotCurSel = true;
    }
  }

  // ----------------------------------------------------------------------------
  // Step 3: Get the customer type list from QuickBooks
  //

  QBCustomerTypeQuery custTypeQuery(m_CompanyFileName,
                                    m_pQbxmlrpWrapper);

  if (!custTypeQuery.GetCustomerTypeListFromQB()) {
    MessageBox(custTypeQuery.GetErrorMsg().c_str(), "Error getting customer type list from QuickBooks");

    // close window
    EndDialog(m_NRetCode);
    PostQuitMessage(0);

    return 0;
  }

  // Status check
  if (custTypeQuery.GetStatusSeverity().substr(0, 5) == "Error") {
    MessageBox(custTypeQuery.GetStatusMessage().c_str(), "Received error when getting customer type list from QuickBooks");

    // close window
    EndDialog(m_NRetCode);
    PostQuitMessage(0);

    return 0;
  }

  // ----------------------------------------------------------------------------
  // Step 4: Populate the customer type combo box.
  //

  QBCustTypeRetVec custTypeRetVec = custTypeQuery.GetCustomerTypeVector();
  QBCustTypeRetVec::iterator jt;

  gotCurSel = false;
  for (jt = custTypeRetVec.begin(); jt != custTypeRetVec.end(); jt++){

    QBCustomerTypeRet *custTypeRet = *jt;

    // Add a customer type to the combo box
    SendDlgItemMessage(IDC_COMBO_TypeList, CB_ADDSTRING, 0, (LPARAM)_T(custTypeRet ->FullName().c_str()));

    // Set the first customer type as the currently selected value
    if (!gotCurSel ) {
      SendDlgItemMessage(IDC_COMBO_TypeList, CB_SETCURSEL, 0, (LPARAM)_T(custTypeRet ->FullName().c_str()));
      gotCurSel = true;
    }
  }

  return 1;
}


/*---------------------------------------------------------------------------
 * OnClickedAddNewCustomer
 * Event handler for the "Add" button
 *
 * It performs the following steps:
 *     1. Collect the form data.
 *     2. If there is a new customer type, add this customer type to
 *        QuickBooks.
 *     3. Add the new customer to QuickBooks.
 *
 */
LRESULT CQBAddCustomerJobDlg::OnClickedAddNewCustomer(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
  // ----------------------------------------------------------------------------
  // Step 1: Collect the form data.
  //
  const int FIELD_SIZE = 256;
  TCHAR companyName[FIELD_SIZE];
  ZeroMemory(companyName, FIELD_SIZE);

  TCHAR name[FIELD_SIZE];
  ZeroMemory(name, FIELD_SIZE);

  TCHAR addCustomerJobFullName[FIELD_SIZE];
  ZeroMemory(addCustomerJobFullName, FIELD_SIZE);

  TCHAR addCustomerJobType[FIELD_SIZE];
  ZeroMemory(addCustomerJobType, FIELD_SIZE);

  TCHAR newCustomerType[FIELD_SIZE];
  ZeroMemory(newCustomerType, FIELD_SIZE);

  TCHAR phone[FIELD_SIZE];
  ZeroMemory(phone, FIELD_SIZE);

  TCHAR email[FIELD_SIZE];
  ZeroMemory(email, FIELD_SIZE);

  // selected Customer Full Name
  this -> GetDlgItemText(IDC_COMBO_CustList, (LPTSTR)addCustomerJobFullName, FIELD_SIZE - 1);

  // selected Customer Type
  this -> GetDlgItemText(IDC_COMBO_TypeList, (LPTSTR)addCustomerJobType, FIELD_SIZE - 1);

  // new Customer Type
  this -> GetDlgItemText(IDC_EDIT_NewType, (LPTSTR)newCustomerType, FIELD_SIZE - 1);

  // Company Name
  this -> GetDlgItemText(IDC_EDIT_CompanyName, (LPTSTR)companyName, FIELD_SIZE - 1);

  // Name
  this -> GetDlgItemText(IDC_EDIT_Name, (LPTSTR)name, FIELD_SIZE - 1);

  // Phone
  this -> GetDlgItemText(IDC_EDIT_Phone, (LPTSTR)phone, FIELD_SIZE - 1);

  // Email
  this -> GetDlgItemText(IDC_EDIT_Email, (LPTSTR)email, FIELD_SIZE - 1);

  // Enforce our required field
  if (name[0] == '\0') {
    MessageBox("Please enter the Customer Name!", _T("Form field is empty"));
    return 0;
  }


  // ----------------------------------------------------------------------------
  // Step 2: If a new customer type was specified, add it to QuickBooks.
  //

  // m_NewCustomer holds the data that will be used to build the
  // CustomerAdd request.
  if (m_NewCustomer != NULL) {
    delete m_NewCustomer;
    m_NewCustomer = NULL;
  }
  m_NewCustomer = new QBCustomerRet;

  if (newCustomerType[0] != '\0') {
    // A new Customer Type was specified, so add it to QuickBooks
    // and use it as the customer type for the new customer
    QBCustomerTypeRet custTypeRet;
    custTypeRet.FullName(newCustomerType);

    // The QBCustomerTypeAdd object will help us send it to QuickBooks.
    QBCustomerTypeAdd custTypeAdd(m_CompanyFileName,
                                  m_pQbxmlrpWrapper);

    custTypeAdd.CustomerTypeRet(&custTypeRet);

    // Add this new Customer Type
    if (!custTypeAdd.AddCustomerType()) {
      MessageBox(custTypeAdd.GetErrorMsg().c_str(), "Error adding new customer type to QuickBooks");
      return 0;
    }

    if (custTypeAdd.GetStatusSeverity().substr(0, 5) == "Error") {
      MessageBox(custTypeAdd.GetStatusMessage().c_str(), "Received error when adding new customer type to QuickBooks");
      return 0;
    }

    m_NewCustomer ->CustomerTypeRef_FullName(newCustomerType);
  }
  else {
    // No new Customer Type was specified, so use the selected
    // Customer Type for the new customer
    m_NewCustomer ->CustomerTypeRef_FullName(addCustomerJobType);
  }


  // ----------------------------------------------------------------------------
  // Step 3: Add the new customer to QuickBooks.
  //

  // Set up the remaining members of this new Customer
  m_NewCustomer ->CompanyName(companyName);
  m_NewCustomer ->Name(name);
  m_NewCustomer ->Phone(phone);
  m_NewCustomer ->Email(email);

  // set the parent ref
  m_NewCustomer ->ParentRef_FullName(addCustomerJobFullName);

  // The QBCustomerAdd object will help us send it to QuickBooks.
  QBCustomerAdd custAdd(m_CompanyFileName,
                        m_pQbxmlrpWrapper);

  custAdd.CustomerRet(m_NewCustomer);

  // Add this customer to QuickBooks
  if (!custAdd.AddCustomer()) { // failed
    MessageBox(custAdd.GetErrorMsg().c_str(), _T("Error adding new customer to QuickBooks"));
  }
  else {
    MessageBox(custAdd.GetStatusMessage().c_str(), _T("Status from QuickBooks when adding new customer"));
  }

  return 0;
}
