// ItemQuery.h : main header file for the ITEMQUERY application
//

#if !defined(AFX_ITEMQUERY_H__ACC985C7_11FB_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_ITEMQUERY_H__ACC985C7_11FB_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CItemQueryApp:
// See ItemQuery.cpp for the implementation of this class
//

class CItemQueryApp : public CWinApp
{
public:
	CItemQueryApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CItemQueryApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CItemQueryApp)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ITEMQUERY_H__ACC985C7_11FB_11D6_AF14_0060088F2F22__INCLUDED_)
