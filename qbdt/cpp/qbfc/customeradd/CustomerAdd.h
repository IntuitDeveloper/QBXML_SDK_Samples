// CustomerAdd.h : main header file for the CUSTOMERADD application
//

#if !defined(AFX_CUSTOMERADD_H__6260F6C8_11E7_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_CUSTOMERADD_H__6260F6C8_11E7_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CCustomerAddApp:
// See CustomerAdd.cpp for the implementation of this class
//

class CCustomerAddApp : public CWinApp
{
public:
	CCustomerAddApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCustomerAddApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CCustomerAddApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CUSTOMERADD_H__6260F6C8_11E7_11D6_AF14_0060088F2F22__INCLUDED_)
