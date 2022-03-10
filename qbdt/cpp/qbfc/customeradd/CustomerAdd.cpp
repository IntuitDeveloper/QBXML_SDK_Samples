// CustomerAdd.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "CustomerAdd.h"
#include "CustomerAddDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CCustomerAddApp

BEGIN_MESSAGE_MAP(CCustomerAddApp, CWinApp)
	//{{AFX_MSG_MAP(CCustomerAddApp)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CCustomerAddApp construction

CCustomerAddApp::CCustomerAddApp()
{
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CCustomerAddApp object

CCustomerAddApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CCustomerAddApp initialization

BOOL CCustomerAddApp::InitInstance()
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

	CCustomerAddDlg dlg;
	m_pMainWnd = &dlg;
	int nResponse = dlg.DoModal();

  ::CoUninitialize ();

  // Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
