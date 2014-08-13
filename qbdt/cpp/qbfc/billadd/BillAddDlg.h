// BillAddDlg.h : header file
//

#if !defined(AFX_BILLADDDLG_H__855DD258_1435_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_BILLADDDLG_H__855DD258_1435_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CBillAddDlg dialog

class CBillAddDlg : public CDialog
{
// Construction
public:
	CBillAddDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CBillAddDlg)
	enum { IDD = IDD_BILLADD_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBillAddDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CBillAddDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.


#endif // !defined(AFX_BILLADDDLG_H__855DD258_1435_11D6_AF14_0060088F2F22__INCLUDED_)

