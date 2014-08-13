/*---------------------------------------------------------------------------
 * FILE: EventSubscribe.cpp
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
#include "EventSubscribe.h"
#include "EventSubscribeDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// MFC-managed message map.
BEGIN_MESSAGE_MAP(CEventSubscribeApp, CWinApp)
	//{{AFX_MSG_MAP(CEventSubscribeApp)
	//}}AFX_MSG_MAP
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()


// Constructor.
CEventSubscribeApp::CEventSubscribeApp()
{
}


// The one and only CEventSubscribeApp object
CEventSubscribeApp theApp;


// CEventSubscribeApp initialization
BOOL CEventSubscribeApp::InitInstance()
{
	CWinApp::InitInstance();

	CoInitialize( NULL );

    // Display main dialog.
	CEventSubscribeDlg dlg;
	m_pMainWnd = &dlg;
	dlg.DoModal();

	CoUninitialize();

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
