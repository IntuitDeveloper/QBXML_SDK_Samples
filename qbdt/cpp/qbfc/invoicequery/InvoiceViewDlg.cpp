// InvoiceViewDlg.cpp : implementation file
//

#include "stdafx.h"
#include "InvoiceQuery.h"
#include "InvoiceViewDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// InvoiceViewDlg dialog


InvoiceViewDlg::InvoiceViewDlg(const char* invoiceList, CWnd* pParent /*=NULL*/)
	: CDialog(InvoiceViewDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(InvoiceViewDlg)
	m_invoiceList = invoiceList;
	//}}AFX_DATA_INIT
}


void InvoiceViewDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(InvoiceViewDlg)
	DDX_Text(pDX, IDC_INVOICELIST, m_invoiceList);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(InvoiceViewDlg, CDialog)
	//{{AFX_MSG_MAP(InvoiceViewDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// InvoiceViewDlg message handlers
