#if !defined(AFX_INVOICEVIEWDLG_H__27996814_113F_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_INVOICEVIEWDLG_H__27996814_113F_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// InvoiceViewDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// InvoiceViewDlg dialog

class InvoiceViewDlg : public CDialog
{
// Construction
public:
	InvoiceViewDlg(const char* invoiceList, CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(InvoiceViewDlg)
	enum { IDD = IDD_INVOICEVIEW_DIALOG };
	CString	m_invoiceList;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(InvoiceViewDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(InvoiceViewDlg)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_INVOICEVIEWDLG_H__27996814_113F_11D6_AF14_0060088F2F22__INCLUDED_)
