/*---------------------------------------------------------------------------
 * FILE: InvokeUIDlg.h
 *
 * Description:
 * Header file of CInvokeUIDlg, MFC-generated dialog class.
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

#include "DOMXMLBuilder.h"
#include "qbXMLRPWrapper.h"

// Forward declaration.
class qbXMLObjectList;

// CInvokeUIDlg dialog
class CInvokeUIDlg : public CDialog
{
public:
    // Constructor.
	CInvokeUIDlg(CWnd* pParent = NULL);
    // Destructor.
    virtual ~CInvokeUIDlg();

    // MFC-managed Dialog Data.
	//{{AFX_DATA(CInvokeUIDlg)
	enum { IDD = IDD_INVOKEUI_DIALOG };
	//}}AFX_DATA

    // MFC-managed virtual method overrides.
	//{{AFX_VIRTUAL(CInvokeUIDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

protected:
	HICON m_hIcon;

    // MFC-managed message map.
	//{{AFX_MSG(CInvokeUIDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnTxnTypeView();
	afx_msg void OnTxnTypeLoad();
	afx_msg void OnTxnListView();
	afx_msg void OnItemTypeView();
	afx_msg void OnItemTypeLoad();
	afx_msg void OnItemListView();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

private:
    void LoadList( UINT nTypeList, UINT nItemList, qbXMLObjectList* pList );
    void OpenAddWindow( UINT nTypeList, LPCTSTR lpszTag );
    void OpenModWindow( UINT nList, 
                        LPCTSTR lpszTag, 
                        LPCTSTR lpszIDTag );

	BOOL DoQueryRequest( LPCTSTR lpszType, string& strResponse );
	string BuildQueryRequest( LPCTSTR lpszType );

	BOOL DoInvokeRequest( LPCTSTR lpszTag,
                         LPCTSTR lpszType,
                         LPCTSTR lpszIDTag,
                         LPCTSTR lpszIDValuestring );
	string BuildInvokeRequest( LPCTSTR lpszTag,
                              LPCTSTR lpszType,
                              LPCTSTR lpszIDTag,
                              LPCTSTR lpszIDValue );

    void DisplayRequestError( qbXMLRPWrapper& requestProcessor,
                              LPCTSTR lpszMessage );
	void DisplayXMLError( DOMXMLBuilder& xmlBuilder );

   	qbXMLObjectList *m_pTxnList;
    qbXMLObjectList *m_pItemList;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
