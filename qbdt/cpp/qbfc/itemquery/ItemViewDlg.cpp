// ItemViewDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ItemQuery.h"
#include "ItemViewDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// ItemViewDlg dialog


ItemViewDlg::ItemViewDlg(CString itemList, CWnd* pParent /*=NULL*/)
	: CDialog(ItemViewDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(ItemViewDlg)
	m_itemList = itemList;
	//}}AFX_DATA_INIT
}


void ItemViewDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(ItemViewDlg)
	DDX_Text(pDX, IDC_ITEMLIST, m_itemList);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(ItemViewDlg, CDialog)
	//{{AFX_MSG_MAP(ItemViewDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// ItemViewDlg message handlers
