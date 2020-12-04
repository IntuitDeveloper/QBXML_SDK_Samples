
/* MultipleRequestsDlg.cpp : implementation file
 *
 *
 * Description:  This file demonstrates the use of QBFC. It adds
 *               multiple requests and sends them in one set.
 *               The responses come in the same order as the requests.
 *
 * Created On: 8/15/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *      http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#include "stdafx.h"
#include "MultipleRequests.h"
#include "MultipleRequestsDlg.h"

CString QBFCLatestVersion(IQBSessionManagerPtr SessionManager);
IMsgSetRequestPtr GetLatestMsgSetRequest(IQBSessionManagerPtr SessionManager);


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMultipleRequestsDlg dialog

CMultipleRequestsDlg::CMultipleRequestsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CMultipleRequestsDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CMultipleRequestsDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMultipleRequestsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CMultipleRequestsDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CMultipleRequestsDlg, CDialog)
	//{{AFX_MSG_MAP(CMultipleRequestsDlg)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMultipleRequestsDlg message handlers

BOOL CMultipleRequestsDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMultipleRequestsDlg::OnPaint() 
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
HCURSOR CMultipleRequestsDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

#define APPID "123"
#define APPNAME "IDN QBFC VC++ MultipleRequests"
#define CUSTOMER_NAME "Jim"
#define INVOICE_REF_NUMBER "12"
#define CUSTOMER_FIRST_NAME "Jimmy"

#define QBXML_MAJORVER 2
#define QBXML_MINORVER 0

void CMultipleRequestsDlg::OnOK() 
{
  try {

    //Create session manager
    IQBSessionManagerPtr sessionManager(CLSID_QBSessionManager);
    sessionManager->OpenConnection ( APPID, APPNAME );
    sessionManager->BeginSession ( "", omDontCare );

    //Create message set request object
    IMsgSetRequestPtr requestMsgSet = GetLatestMsgSetRequest(sessionManager);
    requestMsgSet->Attributes->OnError = roeContinue;
    
    //Add multiple requests to the message set request object

    //First add customer
    ICustomerAddPtr customerAdd = requestMsgSet->AppendCustomerAddRq();
    BuildCustomerAdd ( customerAdd );
    
    //Add an invoice for this customer
    IInvoiceAddPtr invoiceAdd = requestMsgSet->AppendInvoiceAddRq();
    BuildInvoiceAdd ( invoiceAdd );
    
    //Add query request for this invoice
    IInvoiceQueryPtr invoiceQuery = requestMsgSet->AppendInvoiceQueryRq();
    BuildInvoiceQuery( invoiceQuery);
    
    //Perform the requests
    requestMsgSet->Attributes->OnError = roeContinue;
    
    IMsgSetResponsePtr responseMsgSet = sessionManager->DoRequests(requestMsgSet);
    
    //Close connection because we don't need it anymore
    sessionManager->EndSession();
    sessionManager->CloseConnection();
    
    // Uncomment the following to see the request and response XML for debugging
    //_bstr_t rq = requestMsgSet->ToXMLString();
    //_bstr_t rs = responseMsgSet->ToXMLString();
   
    IResponseListPtr rsList = responseMsgSet->ResponseList;
    int Cnt;
    CString msg;
    IResponsePtr response;
    
    //rsList->Count is 3, because we've added 3 requests
    //Let's retrieve them one by one

    //First retrieve customer add response
    Cnt = 0;
    response = rsList->GetAt(Cnt);
    msg = "Response 0 " ;
    msg += InterpretCustomerAddResponse(response);
    
    //Retrieve invoice add response
    Cnt = 1;
    response = rsList->GetAt(Cnt);
    msg += "\r\nResponse 1 " ;
    msg += InterpretInvoiceAddResponse(response);
    
    //Retrieve invoice query response
    Cnt = 2;
    response = rsList->GetAt(Cnt);
    msg += "\r\nResponse 1 " ;
    msg += InterpretInvoiceQueryResponse(response);
    
        
    AfxMessageBox( msg );
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

void CMultipleRequestsDlg::BuildCustomerAdd(ICustomerAddPtr& customerAdd)
{
  customerAdd->Name->SetValue( CUSTOMER_NAME );
  customerAdd->FirstName->SetValue( CUSTOMER_FIRST_NAME );
  COleDateTime oleDate;
  if( oleDate.ParseDateTime( "10/31/2001", VAR_DATEVALUEONLY ) )
  {
    customerAdd->JobStartDate->SetValue( oleDate );
  }
}


void CMultipleRequestsDlg::BuildInvoiceAdd( IInvoiceAddPtr& invoiceAdd)
{    
  invoiceAdd->CustomerRef->FullName->SetValue(CUSTOMER_NAME);
  invoiceAdd->RefNumber->SetValue (INVOICE_REF_NUMBER);
    
  //Add 2 lines to invoice
  invoiceAdd->ORInvoiceLineAddList->Append()->InvoiceLineAdd->Desc->
    SetValue( "Description for line 1");
  invoiceAdd->ORInvoiceLineAddList->Append()->InvoiceLineAdd->Desc->
    SetValue( "Description for line 2" );

}  

void CMultipleRequestsDlg::BuildInvoiceQuery( IInvoiceQueryPtr& invoiceQuery)
{    
  invoiceQuery->IncludeLineItems->SetValue (VARIANT_TRUE);

  //Filter by INVOICE_REF_NUMBER
  invoiceQuery->ORInvoiceQuery->InvoiceFilter->ORRefNumberFilter->
    RefNumberRangeFilter->FromRefNumber->SetValue( INVOICE_REF_NUMBER );
  invoiceQuery->ORInvoiceQuery->InvoiceFilter->ORRefNumberFilter->
    RefNumberRangeFilter->ToRefNumber->SetValue( INVOICE_REF_NUMBER );
}

CString CMultipleRequestsDlg::InterpretCustomerAddResponse(
    IResponsePtr& response)
{
  CString msg;
  msg.Format ( "Status: Code = %d, Message = %S"
    ", Severity = %S", response->StatusCode, 
    (BSTR)response->StatusMessage, (BSTR)response->StatusSeverity );

  ICustomerRetPtr customerRet = response->Detail;
  if (customerRet != NULL )
  {
    msg += ", Detail Type = " ;
    msg += customerRet->Type->GetAsString();
    msg += ":" ;
    msg += customerRet->Name->GetValue();
  }

  return msg;
}

CString CMultipleRequestsDlg::InterpretInvoiceAddResponse(
    IResponsePtr& response ) 
{
  CString msg;
  msg.Format ( "Status: Code = %d, Message = %S"
    ", Severity = %S", response->StatusCode, 
    (BSTR)response->StatusMessage, (BSTR)response->StatusSeverity );

  IInvoiceRetPtr invoiceRet = response->Detail;
  if( invoiceRet != NULL )
  {
    msg += ", Detail Type = ";
    msg += invoiceRet->Type->GetAsString();
    msg += ":" ;
    msg += invoiceRet->TxnID->GetValue();
  }
   
  return msg;
}

CString CMultipleRequestsDlg::InterpretInvoiceQueryResponse(
    IResponsePtr& response ) 
{

  CString msg;
  msg.Format ( "Status: Code = %d, Message = %S"
    ", Severity = %S", response->StatusCode, 
    (BSTR)response->StatusMessage, (BSTR)response->StatusSeverity );
    
  IInvoiceRetListPtr invoiceRetList = response->Detail;
    
  if (invoiceRetList != NULL ) 
  {
    int count = invoiceRetList->Count;
    for( int ndx = 0; ndx < count; ndx++ )
    {
      IInvoiceRetPtr invoiceRet = invoiceRetList->GetAt(ndx);

      CString tmp;
      _bstr_t txnID =  invoiceRet->TxnID->GetValue();
      COleDateTime dateTime( invoiceRet->TxnDate->GetValue() );
      _bstr_t refNumber = invoiceRet->RefNumber->GetValue();

      tmp.Format( "\r\n\tInvoice (%d): TxnID = %S, "
        "TxnDate = %s, RefNumber = %S",
        ndx, (BSTR)txnID, dateTime.Format(), (BSTR)refNumber );
      msg += tmp;
    }
  }
  return msg;
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
