/*---------------------------------------------------------------------------
 * FILE: ShowXMLDlg.h
 *
 * Description:
 * Header file of CShowXMLDlg, MFC-generated dialog class.
 *
 * Displays given XML in a manner that allows it to be copied
 * to the clipboard.
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

#include "Resource.h"

// CShowXMLDlg dialog
class CShowXMLDlg : public CDialog
{
public:
    // Constructor.
	CShowXMLDlg(CWnd* pParent = NULL);

    // Initializers.
    void SetTitle( LPCTSTR lpszTitle )   { m_lpszTitle = lpszTitle; }
    void SetHeader( LPCTSTR lpszHeader ) { m_lpszHeader = lpszHeader; }
    void SetXML( LPCTSTR lpszXML )       { m_lpszXML = lpszXML; }

    // MFC-managed Dialog Data.
	//{{AFX_DATA(CShowXMLDlg)
	enum { IDD = IDD_SHOWXML };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


    // MFC-managed virtual method overrides.
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CShowXMLDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

    // MFC-managed message map.
	//{{AFX_MSG(CShowXMLDlg)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

private:
    LPCTSTR m_lpszTitle;
    LPCTSTR m_lpszHeader;
    LPCTSTR m_lpszXML;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
