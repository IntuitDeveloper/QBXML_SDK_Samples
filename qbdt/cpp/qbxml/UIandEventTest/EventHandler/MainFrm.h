/*---------------------------------------------------------------------------
 * FILE: MainFrm.h
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

#if !defined(AFX_MAINFRM_H__0043CC15_9186_47B0_83D5_9561653E5AA4__INCLUDED_)
#define AFX_MAINFRM_H__0043CC15_9186_47B0_83D5_9561653E5AA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ChildView.h"

class CMainFrame : public CFrameWnd
{
	
public:
	CMainFrame();

    // Custom window class name used by PreCreateWindow()
    // to enable CEventHandlerApp::FirstInstance().
    static LPCTSTR WINDOW_CLASS;

    // Make this application the foreground app.
    // See definition for more info.
    void BringToForeground();

    // Display provided event XML in edit control.
    void ShowEventXML( const string& strXML );

protected: 
	DECLARE_DYNAMIC(CMainFrame)

// Attributes
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMainFrame)
	public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual BOOL OnCmdMsg(UINT nID, int nCode, void* pExtra, AFX_CMDHANDLERINFO* pHandlerInfo);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CMainFrame();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:  // control bar embedded members
	CStatusBar  m_wndStatusBar;
	CChildView    m_wndView;

// Generated message map functions
protected:
	//{{AFX_MSG(CMainFrame)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSetFocus(CWnd *pOldWnd);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MAINFRM_H__0043CC15_9186_47B0_83D5_9561653E5AA4__INCLUDED_)
