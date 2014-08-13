// ItemQuery.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "ItemQuery.h"
#include "ItemQueryDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CItemQueryApp

BEGIN_MESSAGE_MAP(CItemQueryApp, CWinApp)
	//{{AFX_MSG_MAP(CItemQueryApp)
	//}}AFX_MSG
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CItemQueryApp construction

CItemQueryApp::CItemQueryApp()
{
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CItemQueryApp object

CItemQueryApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CItemQueryApp initialization

BOOL CItemQueryApp::InitInstance()
{
	AfxEnableControlContainer();

	// Standard initialization

#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

	CItemQueryDlg dlg;
	m_pMainWnd = &dlg;
  
  ::CoInitialize(0);
	int nResponse = dlg.DoModal();
  ::CoUninitialize();

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
