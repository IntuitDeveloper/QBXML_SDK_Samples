/*---------------------------------------------------------------------------
 * FILE: ShowXMLDlg.cpp
 *
 * Description:
 * MFC-generated dialog class.
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

#include "stdafx.h"
#include "ShowXMLDlg.h"
#include "DOMXMLBuilder.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CShowXMLDlg dialog


// MFC-added contructor.
CShowXMLDlg::CShowXMLDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CShowXMLDlg::IDD, pParent)
{
    // Strings used by this dialog will be managed
    // simply as pointers to values set via the
    // initializers. We assume those values will 
    // remain in scope as long as the dialog does.
    m_lpszTitle = m_lpszHeader = m_lpszXML = NULL;

    // MFC-managed elements.
	//{{AFX_DATA_INIT(CShowXMLDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


// MFC-added data exchange
void CShowXMLDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CShowXMLDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


// MFC-managed message map.
BEGIN_MESSAGE_MAP(CShowXMLDlg, CDialog)
	//{{AFX_MSG_MAP(CShowXMLDlg)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CShowXMLDlg message handlers

// OnInitDialog()
// Dialog initialization.
BOOL CShowXMLDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
    if( m_lpszTitle ) {
        CString strTitle = "SHOW QBSDK XML - ";
        strTitle += m_lpszTitle;
        SetWindowText( strTitle );
    }
    
    if( m_lpszHeader ) {
        SetDlgItemText( IDC_SHOWXML_HEADER, m_lpszHeader );
    }

    if( m_lpszXML ) {
        string strXML = DOMXMLBuilder::FormatXML( m_lpszXML );
        SetDlgItemText( IDC_SHOWXML_EDIT, strXML.c_str() );
    }
	
	return TRUE;
}

