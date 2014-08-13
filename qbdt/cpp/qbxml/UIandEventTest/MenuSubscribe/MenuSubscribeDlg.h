/*---------------------------------------------------------------------------
 * FILE: MenuSubscribeDlg.h
 *
 * Description:
 * Header file of CMenuSubscriberDlg, MFC-generated dialog class.
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

// Forward declarations.
class qbXMLRPWrapper;
class DOMXMLBuilder;

// CMenuSubscribeDlg dialog
class CMenuSubscribeDlg : public CDialog
{
public:
    // Constructor.
	CMenuSubscribeDlg(CWnd* pParent = NULL);

    // MFC-managed Dialog Data.
	//{{AFX_DATA(CMenuSubscribeDlg)
	enum { IDD = IDD_MENUSUBSCRIBE_DIALOG };
	//}}AFX_DATA

    // MFC-managed virtual method overrides.
	//{{AFX_VIRTUAL(CMenuSubscribeDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

protected:
	HICON m_hIcon;

    // MFC-managed message map.
	//{{AFX_MSG(CMenuSubscribeDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnAddSubscription();
	afx_msg void OnClickVisibleIf1();
	afx_msg void OnClickVisibleIf2();
	afx_msg void OnClickVisibleIf3();
	afx_msg void OnClickEnabledIf1();
	afx_msg void OnClickEnabledIf2();
	afx_msg void OnClickEnabledIf3();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

private:
	string GetItemText( int id );

	int    GetMenuCount();
	string GetMenuText( int iMenu );
	BOOL   GetVisibleIf( int iMenu );
	string GetVisibleCondition( int iMenu );
	BOOL   GetEnabledIf( int iMenu );
	string GetEnabledCondition( int iMenu );

	void ChangeButton( int id );

	BOOL DoQueryRequest( qbXMLRPWrapper& requestProcessor, BOOL& alreadyExists );
	BOOL DoDeleteRequest( qbXMLRPWrapper& requestProcessor );
	BOOL DoAddRequest( qbXMLRPWrapper& requestProcessor );

	string BuildQueryRequest();
	string BuildDeleteRequest();
	string BuildAddRequest();

	void DisplayRequestError( TCHAR* lpszMessage, qbXMLRPWrapper& requestProcessor );
	void DisplayXMLError( DOMXMLBuilder& xmlBuilder );
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
