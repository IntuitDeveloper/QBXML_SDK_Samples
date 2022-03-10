/*---------------------------------------------------------------------------
 * FILE: MenuSubscribe.h
 *
 * Description:
 * Header file of CMenuSubscribeApp, the MFC-generated application class.
 *
 * Created On: 03/09/2022
 *
 *
 * Copyright � 2021-2022 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols


// CMenuSubscribeApp:
class CMenuSubscribeApp : public CWinApp
{
public:
    // Constructor.
	CMenuSubscribeApp();

    // MFC-managed virtual method overrides.
	//{{AFX_VIRTUAL(CMenuSubscribeApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

    // MFC-managed message map.
	//{{AFX_MSG(CMenuSubscribeApp)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

extern CMenuSubscribeApp theApp;

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
