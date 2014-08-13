/*---------------------------------------------------------------------------
 * FILE: InvokeUI.cpp
 *
 * Description:
 * MFC-generated application class.
 *
 * Created On: 09/15/2003
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#include "stdafx.h"
#include "InvokeUI.h"
#include "InvokeUIDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// MFC-managed message map.
BEGIN_MESSAGE_MAP(CInvokeUIApp, CWinApp)
	//{{AFX_MSG_MAP(CInvokeUIApp)
	//}}AFX_MSG_MAP
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()


// Constructor.
CInvokeUIApp::CInvokeUIApp()
{
}


// The one and only CInvokeUIApp object
CInvokeUIApp theApp;


// CInvokeUIApp initialization
BOOL CInvokeUIApp::InitInstance()
{
	//AfxOleInit();
	// InitCommonControls() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	//vc6 InitCommonControls();
	Enable3dControlsStatic();	// Call this when linking to MFC statically  //vc6

	CWinApp::InitInstance();

	CoInitialize( NULL );

    // Display main dialog.
	CInvokeUIDlg dlg;
	m_pMainWnd = &dlg;
	dlg.DoModal();

	CoUninitialize();

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
