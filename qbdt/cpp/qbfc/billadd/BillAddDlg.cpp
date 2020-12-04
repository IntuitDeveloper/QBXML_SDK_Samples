
/*********************************************************************************
* BillAddDlg.cpp : implementation file
*
* Description:  This sample demonstrates adding a bill using QBFC.
*               It includes examples of the following:
*                   - Adding a BillAdd request and sending it to QuickBooks
*                   - Appending an element to a list
*                   - Setting an OR object (which contains mutually exclusive elements)
*                   - Setting a Ref object (which refers to another list object)
*                   - Checking a field that is not guaranteed in the response
*                     and obtaining data from it
*
* Created On: 8/15/2002
*
* Copyright © 2002-2013 Intuit Inc. All rights reserved.
* Use is subject to the terms specified at:
*      http://developer.intuit.com/legal/devsite_tos.html
*
* 8/9/2013 Updated to qbfc13.dll
***********************************************************************************/


#include "stdafx.h"
#include <math.h>
#include <stdlib.h>
#include "BillAdd.h"
#include "BillAddDlg.h"

#import "qbFC13.dll" no_namespace, named_guids
CString QBFCLatestVersion(IQBSessionManagerPtr SessionManager);
IMsgSetRequestPtr GetLatestMsgSetRequest(IQBSessionManagerPtr SessionManager);


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBillAddDlg dialog

CBillAddDlg::CBillAddDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CBillAddDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CBillAddDlg)
	//}}AFX_DATA_INIT
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CBillAddDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CBillAddDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CBillAddDlg, CDialog)
	//{{AFX_MSG_MAP(CBillAddDlg)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBillAddDlg message handlers

BOOL CBillAddDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CBillAddDlg::OnPaint() 
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

HCURSOR CBillAddDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}


#define APPID        "123"
#define APPNAME      "IDN QBFC VC BillAdd"

void CBillAddDlg::OnOK() 
{
	try 
  {
	 
    //Create session manager
    IQBSessionManagerPtr sessionManager(CLSID_QBSessionManager);
    sessionManager->OpenConnection ( APPID, APPNAME );
    sessionManager->BeginSession ( "", omDontCare );

    //Create message set request object
    IMsgSetRequestPtr requestMsgSet = GetLatestMsgSetRequest(sessionManager);
    //Initialize the message set request's attribute
    requestMsgSet->Attributes->OnError = roeContinue;
    
    //Add the request to the request set object
    IBillAddPtr billAdd = requestMsgSet->AppendBillAddRq();
    //Set the BillAdd field values
    COleDateTime dt = COleDateTime::GetCurrentTime() ; 
    billAdd->TxnDate->SetValue( dt );
    billAdd->VendorRef->FullName->SetValue( "C.U. Electric" );
    
    //Set the list of item lines for the bill

    //First append to the list of ORItemLineAdd objects
    IORItemLineAddPtr ORItemLineAdd = billAdd->ORItemLineAddList->Append();
    //Set ItemLineAdd from the ORItemLineAdd object
    IItemLineAddPtr itemLineAdd = ORItemLineAdd->ItemLineAdd;
    //Set a reference to an item by using the FullName
    itemLineAdd->ItemRef->FullName->SetValue ( "Subs:Electrical" );
    itemLineAdd->Cost->SetValue( 28.31 );
    
    //Perform the request
    IMsgSetResponsePtr responseSet = sessionManager->DoRequests(requestMsgSet);
     
    // Uncomment the following to see the request and response XML for debugging
    //_bstr_t rq = requestMsgSet->ToXMLString();
    //_bstr_t rs = responseSet->ToXMLString();
   
    //Interpret the response
    IResponsePtr response = responseSet->ResponseList->GetAt(0);
    
    CString msg;
    msg.Format ( "Status: Code = %d, Message = %S"
    ", Severity = %S", response->StatusCode, 
    (BSTR)response->StatusMessage, (BSTR)response->StatusSeverity );
  
    //The Detail property of IResponse object
    //returns a Ret object for Add and Mod requests
    //In this case, the detail object is IBillRet
    
    //For help finding out the detail's type use the following line:
    //_bstr_t strType = response->Detail->Type->GetAsString();
    
    IBillRetPtr billRet = response->Detail;
    if ( billRet != NULL ) 
    {
      //Retrive fields that are guaranteed to be returned in the response
      msg += "\r\nBillRet: EditSequence = " ;
      msg += billRet->EditSequence->GetValue();
      
      msg += ", Amount Due = " ;
      msg += billRet->AmountDue->GetAsString();
        
      //Retrieve a field that is not guaranteed to be returned from QuickBooks
      if (billRet->DueDate != NULL) 
      {
        msg += ", Due Date = " ;
        COleDateTime dueDt ( billRet->DueDate->GetValue() );
        msg += dueDt.Format ();
      }
    }    
     
    sessionManager->EndSession();
    sessionManager->CloseConnection();
    
    AfxMessageBox (msg);

  }
  catch ( _com_error e )
  {
    CString msg;
    _bstr_t desc = e.Description();
    HRESULT hr = e.Error();
    if( hr == 0x800401f0 ) 
    {
      msg = "You need to call ::CoInitialize(0) in your application before "
        "any COM call.";
    }
    else {
      msg.Format( "Error: 0x%x %s" , e.Error(), 
        (desc.length() > 0) ? (const char*)desc : "" );
    }
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
