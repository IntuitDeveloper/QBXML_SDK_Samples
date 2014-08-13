/*---------------------------------------------------------------------------
 * FILE: InvokeUI.h
 *
 * Description:
 * Header file of CInvokeUIApp, the MFC-generated application class.
 *
 * Created On: 09/15/2003
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */
#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols


// CInvokeUIApp
class CInvokeUIApp : public CWinApp
{
public:
    // Constructor.
	CInvokeUIApp();

    // MFC-managed virtual method overrides.
	//{{AFX_VIRTUAL(CInvokeUIApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

    // MFC-managed message map.
	//{{AFX_MSG(CInvokeUIApp)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

extern CInvokeUIApp theApp;