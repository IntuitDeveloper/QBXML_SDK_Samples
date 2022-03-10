// CustomerAddDlg.h : header file
//

#if !defined(AFX_CUSTOMERADDDLG_H__6260F6CA_11E7_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_CUSTOMERADDDLG_H__6260F6CA_11E7_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


/////////////////////////////////////////////////////////////////////////////
// CCustomerAddDlg dialog

class CCustomerAddDlg : public CDialog
{
// Construction
public:
	CCustomerAddDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CCustomerAddDlg)
	enum { IDD = IDD_CUSTOMERADD_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCustomerAddDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CCustomerAddDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.



#endif // !defined(AFX_CUSTOMERADDDLG_H__6260F6CA_11E7_11D6_AF14_0060088F2F22__INCLUDED_)
