/*---------------------------------------------------------------------------
 * FILE: MenuSubscribe.cpp
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
#include "MenuSubscribe.h"
#include "MenuSubscribeDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// MFC-managed message map.
BEGIN_MESSAGE_MAP(CMenuSubscribeApp, CWinApp)
	//{{AFX_MSG_MAP(CMenuSubscribeApp)
	//}}AFX_MSG_MAP
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()


// Constructor.
CMenuSubscribeApp::CMenuSubscribeApp()
{
}


// The one and only CMenuSubscribeApp object
CMenuSubscribeApp theApp;


// CMenuSubscribeApp initialization
BOOL CMenuSubscribeApp::InitInstance()
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
	CMenuSubscribeDlg dlg;
	m_pMainWnd = &dlg;
	dlg.DoModal();

	CoUninitialize();

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
