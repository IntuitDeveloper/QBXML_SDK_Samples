/*---------------------------------------------------------------------------
 * FILE: EventHandler.h
 *
 * Description:
 * Header file of CEventHandlerApp, the MFC-generated application class.
 *
 * Created On: 09/15/2003
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */
#if !defined(AFX_EVENTHANDLER_H__B21C18C7_1A95_41F5_AE7A_87946AF724D8__INCLUDED_)
#define AFX_EVENTHANDLER_H__B21C18C7_1A95_41F5_AE7A_87946AF724D8__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols
#include "EventHandler_i.h"

/////////////////////////////////////////////////////////////////////////////
// CEventHandlerApp:
// See EventHandler.cpp for the implementation of this class
//

class CEventHandlerApp : public CWinApp
{
public:
	CEventHandlerApp();

	static BOOL IsCallbackRegistered();
    
    BOOL FirstInstance();       // Allow only one instance of app.
    void CreateMainWindow();    // Allow creation of main window from QBSDKCallback.

    // Constants for profile settings.
    static LPCTSTR lpszProfileSection;
    static LPCTSTR lpszProfileLaunch;
    static LPCTSTR lpszProfileOutputPath;
    static LPCTSTR lpszProfileRecoveryAlert;

    // MFC-managed virtual method overrides.
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEventHandlerApp)
	public:
	virtual BOOL InitInstance();
		virtual int ExitInstance();
	//}}AFX_VIRTUAL

public:
    // MFC-managed message map.
	//{{AFX_MSG(CEventHandlerApp)
	afx_msg void OnAppAbout();
	afx_msg void OnFileRegister();
	afx_msg void OnFileUnregister();
	afx_msg void OnUpdateFileRegister(CCmdUI* pCmdUI);
	afx_msg void OnUpdateFileUnregister(CCmdUI* pCmdUI);
	afx_msg void OnEditSettings();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// MFC-added ATL stuff.
private:
	BOOL m_bATLInited;
	BOOL InitATL();
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EVENTHANDLER_H__B21C18C7_1A95_41F5_AE7A_87946AF724D8__INCLUDED_)
