// MultipleRequests.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "MultipleRequests.h"
#include "MultipleRequestsDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMultipleRequestsApp

BEGIN_MESSAGE_MAP(CMultipleRequestsApp, CWinApp)
	//{{AFX_MSG_MAP(CMultipleRequestsApp)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMultipleRequestsApp construction

CMultipleRequestsApp::CMultipleRequestsApp()
{
	// Place all significant initialization in InitInstance
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CMultipleRequestsApp object

CMultipleRequestsApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CMultipleRequestsApp initialization

BOOL CMultipleRequestsApp::InitInstance()
{
	AfxEnableControlContainer();

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	//  of your final executable, you should remove from the following
	//  the specific initialization routines you do not need.

#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

  ::CoInitialize(0);

	CMultipleRequestsDlg dlg;
	m_pMainWnd = &dlg;
	int nResponse = dlg.DoModal();

  ::CoUninitialize ();

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
