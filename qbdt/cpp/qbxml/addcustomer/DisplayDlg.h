/*---------------------------------------------------------------------------
 * FILE: DisplayDlg.h
 *
 * Description:
 * ATL Dlg class, displays the XML passed to its constructor.
 *
 * Created On: 11/07/2001
 * Updated to QBXML 2.0 August 2002
 *
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */


#ifndef __DISPLAYDLG_H_
#define __DISPLAYDLG_H_

#include "StdAfx.h"
#include "resource.h"       // main symbols
#include <atlhost.h>

/////////////////////////////////////////////////////////////////////////////
// CDisplayDlg
class CDisplayDlg :
  public CAxDialogImpl<CDisplayDlg>
{
public:
  CDisplayDlg(const string &xml)
    :m_XML(xml)
  {
  }

  ~CDisplayDlg()
  {
  }

  enum { IDD = IDD_DISPLAYDLG };

BEGIN_MSG_MAP(CDisplayDlg)
  MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
  COMMAND_ID_HANDLER(IDOK, OnOK)
  COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
  COMMAND_HANDLER(IDC_DONE, BN_CLICKED, OnClickedDone)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

  LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
  {
    // Set the XML Edit box
    SetDlgItemText(IDC_XML, m_XML.c_str());
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
    EndDialog(wID);
    PostQuitMessage(0);
    return 0;
  }

  LRESULT OnClickedDone(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
  {
    EndDialog(wID);
    PostQuitMessage(0);
    return 0;
  }

private:
  string m_XML;

  // no default ctor is provided
  //
  CDisplayDlg(){};
};

#endif //__DISPLAYDLG_H_
