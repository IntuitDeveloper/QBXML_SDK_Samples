/*---------------------------------------------------------------------------
 * FILE: QBAddCustomerJobDlg.h
 *
 * Description:
 * ATL Dialog class.  This window allows user to add a new Customer job
 * into QB.  It displays a customer list combo box and a customer type
 * combo box to allow user to select an existing parent customer and
 * an existing customer type.  There's also an edit box provided that
 * allows the user to create a new customer type (and associate that
 * new customer type with the new customer job).
 *
 * Created On: 11/11/2001
 * Updated to SDK 2.0: 08/08/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#ifndef __QBAddCustomerJobDLG_H_
#define __QBAddCustomerJobDLG_H_

#include "resource.h"       // main symbols
#include <atlhost.h>

#include "qbXMLRPWrapper.h"
#include "QBCustomerRet.h"


/////////////////////////////////////////////////////////////////////////////
// CQBAddCustomerJobDlg
class CQBAddCustomerJobDlg :
  public CAxDialogImpl<CQBAddCustomerJobDlg>
{
public:
  CQBAddCustomerJobDlg(TCHAR *companyFile, qbXMLRPWrapper *pQbxmlrpWrapper);
  ~CQBAddCustomerJobDlg();

  enum { IDD = IDD_QBADDCUSTOMERJOBDLG };

BEGIN_MSG_MAP(CQBAddCustomerJobDlg)
  MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
  COMMAND_ID_HANDLER(IDOK, OnOK)
  COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
  COMMAND_HANDLER(IDC_AddNewCustomer, BN_CLICKED, OnClickedAddNewCustomer)
  COMMAND_HANDLER(IDC_Exit, BN_CLICKED, OnClickedExit)
END_MSG_MAP()

  // Initialize the Dialog
  LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

  LRESULT OnOK(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
  {
    DestroyWindow();
    PostQuitMessage(0);
    return 0;
  }

  void ShutDown()
  {
    DestroyWindow();
//    PostQuitMessage(0);
  }

  LRESULT OnClickedExit(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
  {
    EndDialog(m_NRetCode);
    PostQuitMessage(0);
    return 0;
  }

  LRESULT OnCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
  {
    //DestroyWindow();
    EndDialog(m_NRetCode);
    PostQuitMessage(0);
    return 0;
  }

  LRESULT OnClickedAddNewCustomer(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

  void SetModalReturnCode(int mCode) { m_NRetCode = mCode;};

private:
  CQBAddCustomerJobDlg(){}

  // members
  string          m_CompanyFileName;
  qbXMLRPWrapper  *m_pQbxmlrpWrapper;

  QBCustomerRet   *m_NewCustomer;

  int m_NRetCode;               // parent return code
};

#endif //__QBAddCustomerJobDLG_H_
