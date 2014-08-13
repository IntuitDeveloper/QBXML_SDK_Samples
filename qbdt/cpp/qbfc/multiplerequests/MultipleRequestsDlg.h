// MultipleRequestsDlg.h : header file
//

#if !defined(AFX_MULTIPLEREQUESTSDLG_H__BF498EEA_11DC_11D6_AF14_0060088F2F22__INCLUDED_)
#define AFX_MULTIPLEREQUESTSDLG_H__BF498EEA_11DC_11D6_AF14_0060088F2F22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CMultipleRequestsDlg dialog
#import "qbFC13.dll" no_namespace, named_guids

class CMultipleRequestsDlg : public CDialog
{
// Construction
public:
	CMultipleRequestsDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CMultipleRequestsDlg)
	enum { IDD = IDD_MULTIPLEREQUESTS_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMultipleRequestsDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CMultipleRequestsDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

protected:
  void BuildCustomerAdd( ICustomerAddPtr& customerAdd );
  void BuildInvoiceAdd( IInvoiceAddPtr& invoiceAdd);
  void BuildInvoiceQuery( IInvoiceQueryPtr& invoiceQuery);
  
  CString InterpretCustomerAddResponse(IResponsePtr& response);
  CString InterpretInvoiceAddResponse(IResponsePtr& response) ;
  CString InterpretInvoiceQueryResponse(IResponsePtr& response) ;

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MULTIPLEREQUESTSDLG_H__BF498EEA_11DC_11D6_AF14_0060088F2F22__INCLUDED_)
