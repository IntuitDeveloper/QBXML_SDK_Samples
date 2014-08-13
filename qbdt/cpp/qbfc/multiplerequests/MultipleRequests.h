// MultipleRequests.h : main header file for the MULTIPLEREQUESTS application
//

#if !defined(AFX_MULTIPLEREQUESTS_H__BF498EE8_11DC_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_MULTIPLEREQUESTS_H__BF498EE8_11DC_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CMultipleRequestsApp:
// See MultipleRequests.cpp for the implementation of this class
//

class CMultipleRequestsApp : public CWinApp
{
public:
	CMultipleRequestsApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMultipleRequestsApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CMultipleRequestsApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MULTIPLEREQUESTS_H__BF498EE8_11DC_11D6_AF14_0060088F2F22__INCLUDED_)
