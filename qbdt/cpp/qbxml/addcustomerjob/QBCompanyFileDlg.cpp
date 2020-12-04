/*---------------------------------------------------------------------------
 * FILE: QBCustomerFileDlg.cpp
 *
 * Description:
 * Event handler implementation.
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
#include "QBCompanyFileDlg.h"
#include "QBCustomerRet.h"
#include "QBCustomerQuery.h"
#include "QBAddCustomerJobDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CQBCompanyFileDlg


/*---------------------------------------------------------------------------
 * Constructor
 *
 */
CQBCompanyFileDlg::CQBCompanyFileDlg()
: m_pQbxmlrpWrapper(NULL)
{

}


/*---------------------------------------------------------------------------
 * Destructor
 *
 */
CQBCompanyFileDlg::~CQBCompanyFileDlg()
{

}


/*---------------------------------------------------------------------------
 * OnClickedNext
 * Reads the chosen company file name and brings up the
 * modal Add Sub Customer dialog window.
 *
 */
LRESULT CQBCompanyFileDlg::OnClickedNext(WORD wNotifyCode,
                                         WORD wID,
                                         HWND hWndCtl,
                                         BOOL& bHandled)
{
  // Collect data from the Dlg
  TCHAR companyFile[256];
  companyFile[0] = '\0';

  // Company File Name
  this -> GetDlgItemText(IDC_CompanyFile, (LPTSTR) companyFile, 255);
/* Can't we just use the currently opened company? [i.e. companyFile == ""
  if (companyFile[0] == '\0') {
    MessageBox("No Company File Name!", _T("Error"));
    return 0;
  }*/

  if (!m_pQbxmlrpWrapper ->OpenCompanyFile(companyFile)) {
    MessageBox(m_pQbxmlrpWrapper ->GetErrorMsg().c_str(), _T("Error"));
    return 0;
  }

  // Create the Add Sub Customer dialog window, use Modal Window
  CQBAddCustomerJobDlg cAddCustomerJobDlg(companyFile, m_pQbxmlrpWrapper);

  int ret = cAddCustomerJobDlg.DoModal(::GetActiveWindow());

  cAddCustomerJobDlg.SetModalReturnCode(ret);

  MSG msg;

  // to make tab work
  while (GetMessage(&msg, 0, 0, 0)) {
    if ((cAddCustomerJobDlg) &&
        (!::IsDialogMessage(cAddCustomerJobDlg.m_hWnd,&msg))) {
      DispatchMessage(&msg);
    }
  }

  return 0;
}


/*---------------------------------------------------------------------------
 * OnClickedBrowse
 * Brings up a standard File Open dialog
 *
 */
LRESULT CQBCompanyFileDlg::OnClickedBrowse(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
  OPENFILENAME of;
  TCHAR        szFile[MAX_PATH];

  GetDlgItemText(IDC_CompanyFile, szFile, MAX_PATH);
  memset(&of, 0, sizeof(OPENFILENAME));

  of.lStructSize  = sizeof(OPENFILENAME);
  of.hwndOwner    = this->m_hWnd;
  of.lpstrFilter  = _T("QuickBooks Data Files (*.QBW)\0*.QBW\0\0");
  of.nFilterIndex = 0L;
  of.lpstrFile    = szFile;
  of.nMaxFile     = MAX_PATH;
  of.lpstrTitle   = "Browse for QuickBooks File";
  of.Flags        =  OFN_SHOWHELP | OFN_PATHMUSTEXIST |
                     OFN_FILEMUSTEXIST | OFN_HIDEREADONLY |
                     OFN_EXPLORER | OFN_LONGNAMES;

  if (GetOpenFileName(&of)){
    SetDlgItemText(IDC_CompanyFile, of.lpstrFile);
  }

  return 0L;
}


/*---------------------------------------------------------------------------
 * InitFileName
 * Sets up path to company file based on the registry.
 *
 */
TCHAR * CQBCompanyFileDlg::InitFileName(TCHAR * szDefaultPath,
                                        TCHAR * szFileName)
{
  HKEY    hKey = NULL;
  static  TCHAR szPath[256];  // make it static so we can return it

  szPath[0] = '\0';

  return szPath;
}


/*---------------------------------------------------------------------------
 * DestroyQbxmlrpWrapper
 * Private helper to free the QBXMLRPWrapper.
 *
 */
void CQBCompanyFileDlg::DestroyQbxmlrpWrapper()
{
  if (m_pQbxmlrpWrapper != NULL) {
    delete m_pQbxmlrpWrapper;
    m_pQbxmlrpWrapper = NULL;
  }
}
