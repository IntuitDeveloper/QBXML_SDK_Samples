/*---------------------------------------------------------------------------
 * FILE: MainFrm.cpp
 *
 * Description:
 * MFC-generated frame class.
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

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

// Constant definition of window class name used by PreCreateWindow()
// to enable CEventHandlerApp::FirstInstance().
LPCTSTR CMainFrame::WINDOW_CLASS = _T("EventHandlerWnd");


/////////////////////////////////////////////////////////////////////////////
// CMainFrame

IMPLEMENT_DYNAMIC(CMainFrame, CFrameWnd)

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWnd)
	//{{AFX_MSG_MAP(CMainFrame)
	ON_WM_CREATE()
	ON_WM_SETFOCUS()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

static UINT indicators[] =
{
	ID_SEPARATOR,           // status line indicator
	ID_INDICATOR_CAPS,
	ID_INDICATOR_NUM,
	ID_INDICATOR_SCRL,
};

/////////////////////////////////////////////////////////////////////////////
// CMainFrame construction/destruction

CMainFrame::CMainFrame()
{
}

CMainFrame::~CMainFrame()
{
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CFrameWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	// create a view to occupy the client area of the frame
	if (!m_wndView.Create(NULL, NULL, AFX_WS_DEFAULT_VIEW,
		CRect(0, 0, 0, 0), this, AFX_IDW_PANE_FIRST, NULL))
	{
		TRACE0("Failed to create view window\n");
		return -1;
	}

	if (!m_wndStatusBar.Create(this) ||
		!m_wndStatusBar.SetIndicators(indicators,
		  sizeof(indicators)/sizeof(UINT)))
	{
		TRACE0("Failed to create status bar\n");
		return -1;      // fail to create
	}

	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	if( !CFrameWnd::PreCreateWindow(cs) )
		return FALSE;

    // Register custom window class to enable CEventHandlerApp::FirstInstance().
    // If custom wnd class is not yet registered, register it using
	// class values from default window class.
	WNDCLASS wndcls;
	if( !GetClassInfo(AfxGetInstanceHandle(), CMainFrame::WINDOW_CLASS, &wndcls) ) {
		if( GetClassInfo(AfxGetInstanceHandle(), cs.lpszClass, &wndcls) ) {
			wndcls.lpszClassName = CMainFrame::WINDOW_CLASS;
			if( !AfxRegisterClass( &wndcls ) ) {
				TRACE( "Error registering custom window class." );
				return FALSE;
			}
		}
		else {
			TRACE( "Default window class not found." );
			return FALSE;
		}
	}

	cs.dwExStyle &= ~WS_EX_CLIENTEDGE;
	
    // Use custom window class name to enable CEventHandlerApp::FirstInstance().
    cs.lpszClass = CMainFrame::WINDOW_CLASS; //AfxRegisterWndClass(0);

	return TRUE;
}

// Make this application the foreground app.
// On Win98 and later, only the current app can make a window the foreground.
// If an app tries to make one of its windows the foreground window while
// another app is in the foreground, Windows will flash the app's taskbar
// icon but will not bring it to the foreground. The code below will get
// around this limitation. See Microsoft Knowledge Base articles #97925 and
// #227043 for additional information.
void CMainFrame::BringToForeground() {
	HWND hCurrentWnd = ::GetForegroundWindow();
	int iMyTID = ::GetCurrentThreadId();
	int iCurrTID = ::GetWindowThreadProcessId( hCurrentWnd, 0 );
	::AttachThreadInput( iMyTID, iCurrTID, TRUE );
	SetForegroundWindow();
	::AttachThreadInput( iMyTID, iCurrTID, FALSE );
}

// Display provided event XML in edit control.
void CMainFrame::ShowEventXML( const string& strXML ) {
    m_wndView.SetDlgItemText( IDC_EVENT_XML, strXML.c_str() );
}

/////////////////////////////////////////////////////////////////////////////
// CMainFrame diagnostics

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CFrameWnd::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CFrameWnd::Dump(dc);
}

#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CMainFrame message handlers
void CMainFrame::OnSetFocus(CWnd* pOldWnd)
{
	// forward focus to the view window
	m_wndView.SetFocus();
}

BOOL CMainFrame::OnCmdMsg(UINT nID, int nCode, void* pExtra, AFX_CMDHANDLERINFO* pHandlerInfo)
{
	// let the view have first crack at the command
	if (m_wndView.OnCmdMsg(nID, nCode, pExtra, pHandlerInfo))
		return TRUE;

	// otherwise, do default handling
	return CFrameWnd::OnCmdMsg(nID, nCode, pExtra, pHandlerInfo);
}

