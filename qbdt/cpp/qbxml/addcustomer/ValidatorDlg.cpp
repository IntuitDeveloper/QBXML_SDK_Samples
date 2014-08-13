/*---------------------------------------------------------------------------
 * FILE: ValidatorDlg.cpp
 *
 * Description:
 * ATL Dlg Class ValidatorDlg Implementation
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
#include "ValidatorDlg.h"
#include "XMLValidator.h"

const int FIELD_SIZE = 256;


/*---------------------------------------------------------------------------
 * OnClickedValidate
 * Validate button event handler. Perform the XML validation and
 * display the result.
 *
 */
LRESULT CValidatorDlg::OnClickedValidate(WORD   wNotifyCode,
                                         WORD   wID,
                                         HWND   hWndCtl,
                                         BOOL   & bHandled)
{
  // get the request XML

  GetDlgItemText(IDC_RequestXML, m_RequestXML, XML_SIZE);
  GetDlgItemText(IDC_Schema, m_SchemaFile, MAX_PATH);

  if (m_SchemaFile[0] == '\0') {
    MessageBox("Please enter the schema file name!", _T("Error"));
    return 0;
  }

  bool result = this->ValidateRequestXML();

  MessageBox(m_Msg.c_str(), "Validator Result");

  return 0;
}


/*---------------------------------------------------------------------------
 * OnClickedBrowse
 * Brings up a standard File Open dialog
 *
 */
LRESULT CValidatorDlg::OnClickedBrowse(WORD   wNotifyCode,
                                       WORD   wID,
                                       HWND   hWndCtl,
                                       BOOL   & bHandled)
{
  OPENFILENAME of;
  TCHAR        szFile[MAX_PATH];

  GetDlgItemText(IDC_Schema, szFile, MAX_PATH);
  memset(&of, 0, sizeof(OPENFILENAME));

  of.lStructSize  = sizeof(OPENFILENAME);
  of.hwndOwner    = this->m_hWnd;
  of.lpstrFilter  = _T("Schema File (*.XSD)\0*.XSD\0\0"); ;
  of.nFilterIndex = 0L;
  of.lpstrFile    = szFile;
  of.nMaxFile     = MAX_PATH;
  of.lpstrTitle   = "Browse for QuickBooks Schema File";
  of.Flags        = OFN_SHOWHELP | OFN_PATHMUSTEXIST |
                    OFN_FILEMUSTEXIST | OFN_HIDEREADONLY |
                    OFN_EXPLORER | OFN_LONGNAMES;

  if (GetOpenFileName(&of)) {
    SetDlgItemText(IDC_Schema, of.lpstrFile);
  }

  return 0L;
}


/*---------------------------------------------------------------------------
 * ValidateXML
 * Use XMLValidator to validate Request XML
 *
 */
bool CValidatorDlg::ValidateRequestXML()
{
  bool bResult;

  // Create an XMLValidator obj
  XMLValidator xmlValidator = XMLValidator();

  // Validate the xml
  bResult = xmlValidator.XMLValidation(string(m_RequestXML),
                                       string(m_SchemaFile));
  m_Msg = xmlValidator.GetErrorMsg();

  return bResult;
}
