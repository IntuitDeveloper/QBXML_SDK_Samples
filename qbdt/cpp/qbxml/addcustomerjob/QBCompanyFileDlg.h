/*---------------------------------------------------------------------------
 * FILE: QBCompanyFileDlg.h
 *
 * Description:
 * ATL Dialog class.  It prompts the user for a QuickBooks company file
 * name and brings up the dialog for adding a sub customer.
 *
 * Created On: 11/11/2001
 * Updated to SDK 2.0: 08/08/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#ifndef __QBCOMPANYFILEDLG_H_
#define __QBCOMPANYFILEDLG_H_

#include "resource.h"       // main symbols
#include <atlhost.h>
#include "QBAddCustomerJobDlg.h"
#include "qbXMLRPWrapper.h"

/////////////////////////////////////////////////////////////////////////////
// CQBCompanyFileDlg
class CQBCompanyFileDlg :
  public CAxDialogImpl<CQBCompanyFileDlg>
{
public:
  CQBCompanyFileDlg();
  ~CQBCompanyFileDlg();

  enum { IDD = IDD_QBCOMPANYFILEDLG };

BEGIN_MSG_MAP(CQBCompanyFileDlg)
  MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
  COMMAND_ID_HANDLER(IDOK, OnOK)
  COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
  COMMAND_HANDLER(IDC_Exit, BN_CLICKED, OnClickedExit)
  COMMAND_HANDLER(IDC_Next, BN_CLICKED, OnClickedNext)
  COMMAND_HANDLER(IDC_Browse, BN_CLICKED, OnClickedBrowse)
END_MSG_MAP()

  LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
  {
    // set the default company file
    SetDlgItemText(IDC_CompanyFile, (LPTSTR) InitFileName("c:\\program files\\intuit\\quickbooks pro", "sample_product-based business.qbw"));

    m_pQbxmlrpWrapper = new qbXMLRPWrapper();

    return 1;  // Let the system set the focus
  }

  LRESULT OnOK(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
  {
    DestroyQbxmlrpWrapper();
    DestroyWindow();
    PostQuitMessage(0);
    return 0;
  }

  LRESULT OnCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
  {
    DestroyQbxmlrpWrapper();
    DestroyWindow();
    PostQuitMessage(0);
    return 0;
  }

  LRESULT OnClickedExit(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
  {
    DestroyQbxmlrpWrapper();
    DestroyWindow();
    PostQuitMessage(0);
    return 0;
  }

  LRESULT OnClickedNext(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

  LRESULT OnClickedBrowse(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

private:
  qbXMLRPWrapper *m_pQbxmlrpWrapper;

  // Create File Path
  TCHAR * InitFileName(TCHAR * szDefaultPath, TCHAR * szFileName);

  void DestroyQbxmlrpWrapper();
};

#endif //__QBCOMPANYFILEDLG_H_
