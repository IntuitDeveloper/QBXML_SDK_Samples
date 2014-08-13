// InvoiceQuery.h : main header file for the INVOICEQUERY application
//

#if !defined(AFX_INVOICEQUERY_H__DCE30137_111F_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_INVOICEQUERY_H__DCE30137_111F_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CInvoiceQueryApp:
// See InvoiceQuery.cpp for the implementation of this class
//

class CInvoiceQueryApp : public CWinApp
{
public:
	CInvoiceQueryApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CInvoiceQueryApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CInvoiceQueryApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_INVOICEQUERY_H__DCE30137_111F_11D6_AF14_0060088F2F22__INCLUDED_)
