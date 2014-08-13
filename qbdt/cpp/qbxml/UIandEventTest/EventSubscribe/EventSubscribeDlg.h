/*---------------------------------------------------------------------------
 * FILE: EventSubscribeDlg.h
 *
 * Description:
 * Header file of CEventSubscriberDlg, MFC-generated dialog class.
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

// CEventSubscribeDlg dialog
class CEventSubscribeDlg : public CDialog
{
public:
    // Constructor.
	CEventSubscribeDlg(CWnd* pParent = NULL);

    // MFC-managed Dialog Data.
	//{{AFX_DATA(CEventSubscribeDlg)
	enum { IDD = IDD_EVENTSUBSCRIBE_DIALOG };
	//}}AFX_DATA

    // MFC-managed virtual method overrides.
	//{{AFX_VIRTUAL(CEventSubscribeDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

protected:
	HICON m_hIcon;

    // MFC-managed message map.
	//{{AFX_MSG(CEventSubscribeDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnAddSubscription();
	afx_msg void OnTxnSubscribe();
	afx_msg void OnListSubscribe();
	afx_msg void OnRemove();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

private:
    void AddEvent( const string& strType,
                   UINT  nItemList,
                   UINT  nAdd,
                   UINT  nMod,
                   UINT  nDel,
                   UINT  nMerge );
    void GetEventInfo( const string& strText,
                       string& strSubTag, 
                       string& strTypeTag, 
                       string& strOpTag,
                       string& strItem,
                       BOOL& bAdd, 
                       BOOL& bDel, 
                       BOOL& bMod, 
                       BOOL& bMerge );

	BOOL DoQueryRequest( qbXMLRPWrapper& requestProcessor, BOOL& alreadyExists );
	BOOL DoDeleteRequest( qbXMLRPWrapper& requestProcessor );
	BOOL DoAddRequest( qbXMLRPWrapper& requestProcessor );

	string BuildQueryRequest();
	string BuildDeleteRequest();
	string BuildAddRequest();

    void BuildUIEventAddRequest( DOMXMLBuilder* pXMLBuilder );
    void BuildDataEventAddRequest( DOMXMLBuilder* pXMLBuilder );

    BOOL HasUIEventSubscriptions();
    BOOL HasDataEventSubscriptions();

	void DisplayRequestError( TCHAR* lpszMessage, qbXMLRPWrapper& requestProcessor );
	void DisplayXMLError( DOMXMLBuilder& xmlBuilder );
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
