// BillAdd.h : main header file for the BILLADD application
//

#if !defined(AFX_BILLADD_H__855DD256_1435_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_BILLADD_H__855DD256_1435_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CBillAddApp:
// See BillAdd.cpp for the implementation of this class
//

class CBillAddApp : public CWinApp
{
public:
	CBillAddApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBillAddApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CBillAddApp)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BILLADD_H__855DD256_1435_11D6_AF14_0060088F2F22__INCLUDED_)
