/*---------------------------------------------------------------------------
 * FILE: SettingsDialog.h
 *
 * Description:
 * Header file of CSettingsDialog, MFC-generated dialog class.
 *
 * Allows setting of application options.
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

// CSettingsDialog dialog
class CSettingsDialog : public CDialog
{
public:
    // Constructor.
	CSettingsDialog(CWnd* pParent = NULL);

    // MFC-managed Dialog Data.
	//{{AFX_DATA(CSettingsDialog)
	enum { IDD = IDD_SETTINGS };
	CString	m_strOutputPath;
    int     m_nLaunch;
    int     m_nRecoveryAlert;
	//}}AFX_DATA

    // MFC-managed virtual method overrides.
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSettingsDialog)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
    virtual void OnOK();
	//}}AFX_VIRTUAL

protected:

    // MFC-managed message map.
	// Generated message map functions
	//{{AFX_MSG(CSettingsDialog)
	virtual BOOL OnInitDialog();
	afx_msg void OnBrowse();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

