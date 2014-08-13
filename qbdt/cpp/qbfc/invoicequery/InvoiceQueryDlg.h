// InvoiceQueryDlg.h : header file
//

#if !defined(AFX_INVOICEQUERYDLG_H__DCE30139_111F_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_INVOICEQUERYDLG_H__DCE30139_111F_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CInvoiceQueryDlg dialog

class CInvoiceQueryDlg : public CDialog
{
// Construction
public:
	CInvoiceQueryDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CInvoiceQueryDlg)
	enum { IDD = IDD_INVOICEQUERY_DIALOG };
	BOOL	m_includeLineItems;
	CString	m_fromTxnDate;
	CString	m_toTxnDate;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CInvoiceQueryDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CInvoiceQueryDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

protected:
  static void QueryInvoices( COleDateTime* oleFromTxnDate, 
    COleDateTime* oleToTxnDate, BOOL m_includeLineItems );

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_INVOICEQUERYDLG_H__DCE30139_111F_11D6_AF14_0060088F2F22__INCLUDED_)
