#if !defined(AFX_ITEMVIEWDLG_H__ACC985D3_11FB_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_ITEMVIEWDLG_H__ACC985D3_11FB_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ItemViewDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// ItemViewDlg dialog

class ItemViewDlg : public CDialog
{
// Construction
public:
	ItemViewDlg(CString itemList, CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(ItemViewDlg)
	enum { IDD = IDD_VIEWITEMS_DIALOG };
	CString	m_itemList;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(ItemViewDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(ItemViewDlg)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ITEMVIEWDLG_H__ACC985D3_11FB_11D6_AF14_0060088F2F22__INCLUDED_)
