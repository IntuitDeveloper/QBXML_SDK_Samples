/*---------------------------------------------------------------------------
 * FILE: EventHandler.cpp
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
#include "EventHandler.h"

#include "MainFrm.h"
#include <initguid.h>
#include "EventHandler_i.c"
#include "QBSDKCallback.h"
#include "SettingsDialog.h"
#include "AboutDialog.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

// Static constants.
LPCTSTR CEventHandlerApp::lpszProfileSection        = _T("Settings");
LPCTSTR CEventHandlerApp::lpszProfileLaunch         = _T("Launch");
LPCTSTR CEventHandlerApp::lpszProfileOutputPath     = _T("OutputPath");
LPCTSTR CEventHandlerApp::lpszProfileRecoveryAlert  = _T("RecoveryAlert");


// MFC-managed message map.
BEGIN_MESSAGE_MAP(CEventHandlerApp, CWinApp)
	//{{AFX_MSG_MAP(CEventHandlerApp)
	ON_COMMAND(ID_APP_ABOUT, OnAppAbout)
	ON_COMMAND(ID_FILE_REGISTER, OnFileRegister)
	ON_COMMAND(ID_FILE_UNREGISTER, OnFileUnregister)
	ON_UPDATE_COMMAND_UI(ID_FILE_REGISTER, OnUpdateFileRegister)
	ON_UPDATE_COMMAND_UI(ID_FILE_UNREGISTER, OnUpdateFileUnregister)
	ON_COMMAND(ID_EDIT_SETTINGS, OnEditSettings)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Static member functions.

// IsCallbackRegistered()
// Determine if callback interface is already registered.
// Used by Register/Unregister menu item handlers.
BOOL CEventHandlerApp::IsCallbackRegistered() {
	LPOLESTR lpszProgID = NULL;
	HRESULT hr = ProgIDFromCLSID( CLSID_QBSDKCallback, &lpszProgID );
	if( lpszProgID ) {
		CoTaskMemFree( lpszProgID );
		lpszProgID = NULL;
	}

	return SUCCEEDED(hr);
}


/////////////////////////////////////////////////////////////////////////////
// The one and only CEventHandlerApp object
CEventHandlerApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CEventHandlerApp initialization
CEventHandlerApp::CEventHandlerApp()
{
}

BOOL CEventHandlerApp::InitInstance()
{
    // If app is already running, show it and bail out.
    if( !FirstInstance() ) {
		return false;
	}

    // ATL stuff added by MFC.
	if (!InitATL())
		return FALSE;


	CCommandLineInfo cmdInfo;
	ParseCommandLine(cmdInfo);

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	//  of your final executable, you should remove from the following
	//  the specific initialization routines you do not need.

#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

	// Change the registry key under which our settings are stored.
	SetRegistryKey(_T("QBSDK Tools"));


	if (cmdInfo.m_bRunEmbedded || cmdInfo.m_bRunAutomated)
	{
		return TRUE;
	}


    CreateMainWindow();

    return TRUE;
}


// CreateMainWindow()
// Separate creation of main window so that it can be called
// separately from QBSDKCallback::inform() method.
void CEventHandlerApp::CreateMainWindow()
{
	// To create the main window, this code creates a new frame window
	// object and then sets it as the application's main window object.

	CMainFrame* pFrame = new CMainFrame;
	m_pMainWnd = pFrame;

	// create and load the frame with its resources
	pFrame->LoadFrame(IDR_MAINFRAME,
		WS_OVERLAPPEDWINDOW | FWS_ADDTOTITLE, NULL,
		NULL);

	// The one and only window has been initialized, so show and update it.
	pFrame->ShowWindow(SW_SHOW);
	pFrame->UpdateWindow();
}


/////////////////////////////////////////////////////////////////////////////
// ATL stuff added by MFC.

CEventHandlerModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
OBJECT_ENTRY(CLSID_QBSDKCallback, CQBSDKCallback)
END_OBJECT_MAP()

LONG CEventHandlerModule::Unlock()
{
	AfxOleUnlockApp();
	return 0;
}

LONG CEventHandlerModule::Lock()
{
	AfxOleLockApp();
	return 1;
}

LPCTSTR CEventHandlerModule::FindOneOf(LPCTSTR p1, LPCTSTR p2)
{
	while (*p1 != NULL)
	{
		LPCTSTR p = p2;
		while (*p != NULL)
		{
			if (*p1 == *p)
				return CharNext(p1);
			p = CharNext(p);
		}
		p1++;
	}
	return NULL;
}

int CEventHandlerApp::ExitInstance()
{
	if (m_bATLInited)
	{
		_Module.RevokeClassObjects();
		_Module.Term();
		CoUninitialize();
	}

	return CWinApp::ExitInstance();

}

BOOL CEventHandlerApp::InitATL()
{
	m_bATLInited = TRUE;

#if _WIN32_WINNT >= 0x0400
	HRESULT hRes = CoInitializeEx(NULL, COINIT_MULTITHREADED);
#else
	HRESULT hRes = CoInitialize(NULL);
#endif

	if (FAILED(hRes))
	{
		m_bATLInited = FALSE;
		return FALSE;
	}

	_Module.Init(ObjectMap, AfxGetInstanceHandle());
	_Module.dwThreadID = GetCurrentThreadId();

	LPTSTR lpCmdLine = GetCommandLine(); //this line necessary for _ATL_MIN_CRT
	TCHAR szTokens[] = _T("-/");

	BOOL bRun = TRUE;
	LPCTSTR lpszToken = _Module.FindOneOf(lpCmdLine, szTokens);
	while (lpszToken != NULL)
	{
		if (lstrcmpi(lpszToken, _T("UnregServer"))==0)
		{
			_Module.UpdateRegistryFromResource(IDR_EVENTHANDLER, FALSE);
			_Module.UnregisterServer(TRUE); //TRUE means typelib is unreg'd
			bRun = FALSE;
			break;
		}
		if (lstrcmpi(lpszToken, _T("RegServer"))==0)
		{
			_Module.UpdateRegistryFromResource(IDR_EVENTHANDLER, TRUE);
			_Module.RegisterServer(TRUE);
			bRun = FALSE;
			break;
		}
		lpszToken = _Module.FindOneOf(lpszToken, szTokens);
	}

	if (!bRun)
	{
		m_bATLInited = FALSE;
		_Module.Term();
		CoUninitialize();
		return FALSE;
	}

	hRes = _Module.RegisterClassObjects(CLSCTX_LOCAL_SERVER, 
		REGCLS_MULTIPLEUSE);
	if (FAILED(hRes))
	{
		m_bATLInited = FALSE;
		CoUninitialize();
		return FALSE;
	}	

	return TRUE;

}


/////////////////////////////////////////////////////////////////////////////
// Helper methods.

// FirstInstance()
// Allow only one instance of app to be running.
// See Microsoft Knowledge Base Article #141752 for more info.
// Note that if the app is initially started via COM, no window may
// be created. In that event, a subsequent launch of the app
// will not have a window to find and will therefore launch a new
// instance of the app.
//
BOOL CEventHandlerApp::FirstInstance() 
{
	CWnd *pWndPrev, *pWndChild;

	// Determine if another window with your class name exists...
	if (pWndPrev = CWnd::FindWindow( CMainFrame::WINDOW_CLASS, NULL ) ) {
		// If so, does it have any popups?
		pWndChild = pWndPrev->GetLastActivePopup();

		// If iconic, restore the main window
		if (pWndPrev->IsIconic())
			pWndPrev->ShowWindow(SW_RESTORE);

		// Bring the main window or its popup to
		// the foreground
		pWndChild->SetForegroundWindow();

		// and you are done activating the previous one.
		return FALSE;
	}
	// First instance. Proceed as normal.
	else {
		return TRUE;
	}
}


/////////////////////////////////////////////////////////////////////////////
// CEventHandlerApp message handlers
	
// OnFileRegister()
// Handles menu item click to register the COM callback interface.
void CEventHandlerApp::OnFileRegister() 
{
	_Module.UpdateRegistryFromResource(IDR_EVENTHANDLER, TRUE);
	_Module.RegisterServer(FALSE);
	AfxMessageBox( "COM Callback registered.\n\nIf QuickBooks is currently running, restart it to pick up the new registration." );
}

// OnFileUnregister()
// Handles menu item click to unregister the COM callback interface.
// Useful for testing error message in QuickBooks.
void CEventHandlerApp::OnFileUnregister() 
{
	_Module.UpdateRegistryFromResource(IDR_EVENTHANDLER, FALSE);
    _Module.UnregisterServer(TRUE); //TRUE means typelib is unreg'd
	AfxMessageBox( "COMCallback unregistered.\n\nIn order for QuickBooks to recoginize this change, be sure to:\n   - Exit this application.\n   - Restart QuickBooks if it is running.\n   - Delete any running instance of EventHandler.exe from the Windows Task Manager." );
}

// OnUpdateFileRegister()
// Enable/Disable File->Register menu depending on whether or
// not callback is already registered.
void CEventHandlerApp::OnUpdateFileRegister(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable( !IsCallbackRegistered() );
}

// OnUpdateFileUnregister()
// Enable/Disable File->Unregister menu depending on whether or
// not callback is already registered.
void CEventHandlerApp::OnUpdateFileUnregister(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable( IsCallbackRegistered() );
}

// OnEditSettings()
// Display settings dialog.
void CEventHandlerApp::OnEditSettings() 
{
	CSettingsDialog dlgSettings;
    dlgSettings.DoModal();
}

// Display About box.
void CEventHandlerApp::OnAppAbout()
{
	CAboutDialog aboutDialog;
	aboutDialog.DoModal();
}

