
/*-----------------------------------------------------------
 * CustomerAddDlg.cpp : implementation file
 *
 * Description:  This sample demonstrates the simple use of QBFC,
 *               by adding a new customer to QuickBooks.
 *
 * Created On: 8/15/2002
 *
 * Copyright © 2021-2022 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *      http://developer.intuit.com/legal/devsite_tos.html
 *
 * updated to qbfc15
 *----------------------------------------------------------
 */ 


#include "stdafx.h"
#include "CustomerAdd.h"
#include "CustomerAddDlg.h"

#import "QBFC15.dll" no_namespace, named_guids
CString QBFCLatestVersion(IQBSessionManagerPtr SessionManager);
IMsgSetRequestPtr GetLatestMsgSetRequest(IQBSessionManagerPtr SessionManager);

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CCustomerAddDlg dialog

CCustomerAddDlg::CCustomerAddDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CCustomerAddDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CCustomerAddDlg)
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CCustomerAddDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CCustomerAddDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CCustomerAddDlg, CDialog)
	//{{AFX_MSG_MAP(CCustomerAddDlg)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CCustomerAddDlg message handlers

BOOL CCustomerAddDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CCustomerAddDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CCustomerAddDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

#define APPID        "123"
#define APPNAME      "IDN QBFC VC AddCustomer"
#define CUSTOMERNAME "Jim"

void CCustomerAddDlg::OnOK() 
{
	try {

    //Create session manager
    IQBSessionManagerPtr sessionManager(CLSID_QBSessionManager);
    sessionManager->OpenConnection ( APPID, APPNAME );
    sessionManager->BeginSession ( "", omDontCare );

    //Create the message set request object
    IMsgSetRequestPtr requestMsgSet = GetLatestMsgSetRequest(sessionManager);
    requestMsgSet->Attributes->OnError = roeContinue;
    
    //Add the request to the message set request object
    ICustomerAddPtr customerAdd = requestMsgSet->AppendCustomerAddRq();
    customerAdd->Name->SetValue( CUSTOMERNAME );
    
    //Perform the request
    IMsgSetResponsePtr responseSet = sessionManager->DoRequests(requestMsgSet);
  
    //Interpret the response
    IResponsePtr response = responseSet->ResponseList->GetAt(0);
    
    CString msg;
    if( response->StatusCode != 0 )
    {
      msg.Format ( "Status: Code = %d, Message = %S"
      ", Severity = %S", response->StatusCode, 
      (BSTR)response->StatusMessage, (BSTR)response->StatusSeverity );
        
      AfxMessageBox( msg );
    }

    //The response detail for Add and Mod requests is a 'Ret' object
    //In our case, it's ICustomerRet
    ICustomerRetPtr customerRet = response->Detail;
    if (customerRet != NULL)
    {
      msg = "\r\nDetail Type = ";
      msg += (LPCTSTR) customerRet->Type->GetAsString();
      msg += "; ListID = " ;
      msg += (LPCTSTR) customerRet->ListID->GetValue();
    
      AfxMessageBox( msg );
    }
    
    sessionManager->EndSession();
    sessionManager->CloseConnection();
        

  }
  catch ( _com_error e )
  {
    CString msg;
    _bstr_t desc = e.Description();
    msg.Format( "Error: 0x%x %s" , e.Error(), 
      (desc.length() > 0) ? (const char*)desc : "" );
    AfxMessageBox( msg );
  }

  EndDialog(0);
}

CString QBFCLatestVersion(IQBSessionManagerPtr SessionManager)
{	
    /*
      * Use oldest version to ensure that we work with any QuickBooks (US)
      */
    IMsgSetRequestPtr requestMsgSet = SessionManager->CreateMsgSetRequest("US",1,0);
    requestMsgSet->AppendHostQueryRq();
    
    IMsgSetResponsePtr QueryResponse = SessionManager->DoRequests(requestMsgSet);
    
    /*
      * The response list contains only one response,
      * which corresponds to our single HostQuery request
      */
    IResponsePtr response = QueryResponse->ResponseList->GetAt(0);
    IHostRetPtr HostResponse = response->Detail;
    IBSTRListPtr supportedVersions = HostResponse->SupportedQBXMLVersionList;
    
    long i;
    double vers, LastVers;
    LastVers = 0.0;
    CString retVal;
    for (i=0; i<supportedVersions->Count; i++) {
        vers = atof(supportedVersions->GetAt(i));
        if (vers > LastVers) {
            LastVers = vers;
            retVal = CString((char *)supportedVersions->GetAt(i));
        }
    }
    return retVal;
}

IMsgSetRequestPtr GetLatestMsgSetRequest(IQBSessionManagerPtr SessionManager)
{
    double supportedVersion = atof(QBFCLatestVersion(SessionManager));

    if (supportedVersion >= 2.0) {
        return SessionManager->CreateMsgSetRequest("US",2, 0);
    } else if (supportedVersion == 1.1) {
        return SessionManager->CreateMsgSetRequest("US",1, 1);
    } else {
        return SessionManager->CreateMsgSetRequest("US",1, 0);
    }
}
