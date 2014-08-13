/*---------------------------------------------------------------------------
 * FILE: ValidatorDlg.h
 *
 * Description:
 * ATL Dlg class. It prompts for the request XML, and asks user to enter
 * schema file.  It uses MSXML to validate the request.
 *
 * Created On: 11/07/2001
 * Updated to QBXML 2.0 August 2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#ifndef __VALIDATORDLG_H_
#define __VALIDATORDLG_H_

#include "resource.h"       // main symbols
#include <atlhost.h>

const int XML_SIZE = 5000;
/////////////////////////////////////////////////////////////////////////////
// CValidatorDlg
class CValidatorDlg :
  public CAxDialogImpl<CValidatorDlg>
{
public:
  CValidatorDlg()
  {

  }

  CValidatorDlg(const string &requestXML)
  {
    m_RequestXML[0] = '\0';

    if (requestXML.empty()){
      m_Msg = "Empty XML!";
      return;
    }

    strcpy(m_RequestXML, requestXML.c_str());
    m_Msg = "Validated";

  }

  ~CValidatorDlg()
  {
  }

  enum { IDD = IDD_VALIDATORDLG };

BEGIN_MSG_MAP(CValidatorDlg)
  MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
  COMMAND_ID_HANDLER(IDOK, OnOK)
  COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
  COMMAND_HANDLER(IDC_Exit, BN_CLICKED, OnClickedExit)
  COMMAND_HANDLER(IDC_VALIDATE, BN_CLICKED, OnClickedValidate)
  COMMAND_HANDLER(IDC_Browse, BN_CLICKED, OnClickedBrowse)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

  LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
  {
    SetDlgItemText(IDC_RequestXML, m_RequestXML);
    return 1;  // Let the system set the focus
  }

  LRESULT OnOK(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
  {
    EndDialog(wID);
    PostQuitMessage(0);
    return 0;
  }

  LRESULT OnCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
  {
    // this is a modal, must use EndDialog() method
    //
    EndDialog(m_NRetCode);
    PostQuitMessage(0);
    return 0;
  }
  LRESULT OnClickedExit(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
  {
    // this is a modal, must use EndDialog() method
    //
    EndDialog(m_NRetCode);
    PostQuitMessage(0);
    return 0;
  }

  LRESULT OnClickedValidate(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

  LRESULT OnClickedBrowse(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

  void SetModalReturnCode(int code){m_NRetCode = code;};

private:

  TCHAR  m_RequestXML[XML_SIZE];
  TCHAR  m_SchemaFile[MAX_PATH];

  string m_Msg;

  int    m_NRetCode;  // modal return code, used for EndDialog() call

  bool ValidateRequestXML();
};

#endif //__VALIDATORDLG_H_
