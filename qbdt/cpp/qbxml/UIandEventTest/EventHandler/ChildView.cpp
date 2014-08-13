/*---------------------------------------------------------------------------
 * FILE: ChildView.cpp
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

#include "stdafx.h"
#include "EventHandler.h"
#include "ChildView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CChildView

CChildView::CChildView() :
CFormView( IDD_MAIN_FORM )
{
}

CChildView::~CChildView()
{
}


// MFC-managed message map.
BEGIN_MESSAGE_MAP(CChildView,CFormView )
	//{{AFX_MSG_MAP(CChildView)
	ON_WM_SIZE()
	ON_COMMAND(ID_EDIT_COPY, OnEditCopy)
	ON_UPDATE_COMMAND_UI(ID_EDIT_COPY, OnUpdateEditCopy)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CChildView message handlers

BOOL CChildView::PreCreateWindow(CREATESTRUCT& cs) 
{
	if (!CFormView::PreCreateWindow(cs))
		return FALSE;

	cs.dwExStyle |= WS_EX_CLIENTEDGE;
	cs.style &= ~WS_BORDER;
	cs.lpszClass = AfxRegisterWndClass(CS_HREDRAW|CS_VREDRAW|CS_DBLCLKS, 
		::LoadCursor(NULL, IDC_ARROW), HBRUSH(COLOR_WINDOW+1), NULL);

	return TRUE;
}

BOOL CChildView::Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName,
	DWORD dwRequestedStyle, const RECT& rect, CWnd* pParentWnd, UINT nID,
    CCreateContext* pContext) 
{
    BOOL bReturn = CFormView::Create( lpszClassName, lpszWindowName,
                                      dwRequestedStyle, rect, pParentWnd, 
                                      nID, pContext );

	CRect rectTemplate;
	GetWindowRect(rectTemplate);
	
    int nMapMode;
    CSize sizeTotal;
    CSize sizePage;
    CSize sizeLine;
    GetDeviceScrollSizes( nMapMode, sizeTotal, sizePage, sizeLine );

    CSize sizeView = GetTotalSize();
    
    // Disable scroll bars in view.
    CSize size(0,0);
    SetScrollSizes( MM_TEXT, size );

    return bReturn;
}

void CChildView::PostNcDestroy() 
{
}

// Copy selected XML to clipboard.
void CChildView::OnEditCopy() 
{
    CEdit* pEdit = (CEdit*)GetDlgItem( IDC_EVENT_XML );
    if( pEdit ) {
        pEdit->Copy();
    }
}

// Enables/Disables Copy menu item based on whether text is selected.
void CChildView::OnUpdateEditCopy(CCmdUI* pCmdUI) 
{
    CEdit* pEdit = (CEdit*)GetDlgItem( IDC_EVENT_XML );
    if( pEdit ) {
        int nBegin, nEnd;
        pEdit->GetSel( nBegin, nEnd );
        pCmdUI->Enable( nBegin != nEnd );
    }
}

// Handle sizing of main window by making sure dialog form
// fills the client area.
void CChildView::OnSize(UINT nType, int cx, int cy) 
{
	CFormView::OnSize(nType, cx, cy);
	
    // Get size of child frame.
    CRect rcChildFrame;
    GetClientRect( &rcChildFrame );

	ResizeDialog( &rcChildFrame );
}

// Change size of dialog so that XML display field fills
// the client area of the window.
void CChildView::ResizeDialog( LPRECT pRect ) 
{
    // Get position of header text field.
    CRect rcHeader( 0, 0, 0, 0);
    CWnd* pHeader = GetDlgItem( IDC_HEADER );
    if( pHeader ) {
        pHeader->GetWindowRect( &rcHeader );
        ScreenToClient( &rcHeader );
    }

    // Get position of XML display field.
    CRect rcXML( 0, 0, 0, 0 );
    CWnd* pXML = GetDlgItem( IDC_EVENT_XML );
    if( pXML ) {
        pXML->GetWindowRect( &rcXML );
        ScreenToClient( &rcXML );

        // Get size of child frame.
        CRect rcChildFrame;
        //GetClientRect( &rcChildFrame );
        rcChildFrame = *pRect;

        // Calculate new position and size for XML display field.
        int nVMargin = rcHeader.top;
        int nHMargin = rcXML.left;
        int nHeaderHeight = rcXML.top - rcHeader.top;
        int nXMLHeight = rcChildFrame.Height() - nHeaderHeight - 2*nVMargin;
        int nXMLWidth  = rcChildFrame.Width() - 2*nHMargin;
        rcXML.bottom = rcXML.top + nXMLHeight;
        rcXML.right = rcXML.left + nXMLWidth;

        pXML->MoveWindow( &rcXML, TRUE );
    }
}

