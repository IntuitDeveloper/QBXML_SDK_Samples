/*---------------------------------------------------------------------------
 * FILE: ChildView.h
 *
 * Description:
 * MFC-generated view class.
 *
 * Created On: 09/15/2003
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#if !defined(AFX_CHILDVIEW_H__51283DBF_1581_4A1D_A8DE_6E1F78C8ED7D__INCLUDED_)
#define AFX_CHILDVIEW_H__51283DBF_1581_4A1D_A8DE_6E1F78C8ED7D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CChildView window

class CChildView : public CFormView
{
// Construction
public:
	CChildView();

    BOOL Create( LPCTSTR lpszClassName, 
                 LPCTSTR lpszWindowName,
	             DWORD dwRequestedStyle, 
                 const RECT& rect, 
                 CWnd* pParentWnd, 
                 UINT nID,
	             CCreateContext* pContext);
    virtual void PostNcDestroy();


// Attributes
public:

// Operations
public:
    void ResizeDialog( LPRECT pRect );

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CChildView)
	protected:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CChildView();

	// Generated message map functions
protected:
	//{{AFX_MSG(CChildView)
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnEditCopy();
	afx_msg void OnUpdateEditCopy(CCmdUI* pCmdUI);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CHILDVIEW_H__51283DBF_1581_4A1D_A8DE_6E1F78C8ED7D__INCLUDED_)
