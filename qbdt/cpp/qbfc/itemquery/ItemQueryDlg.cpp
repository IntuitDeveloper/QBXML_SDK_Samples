/*-----------------------------------------------------------
 * ItemQueryDlg.cpp : implementation file
 *
 * Description:  This module demonstrates querying items using QBFC.
 *               Querying items is similar to querying terms or entities
 *               in the sense that the returned detail object is a list of
 *               'OR' objects  - ORItemRet in this case.
 *               It includes examples of the following:
 *                   - Constructing a query
 *                   - Looping through the list of items in a response
 *                   - Reading data from an OR object
 *                   - Obtaining returned fields, also checking for
 *                     fields that may not exist in the response
 *
 * Created On: 8/15/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *      http://developer.intuit.com/legal/devsite_tos.html
 *
 * 8/9/2013 updated to qbfc13
 *----------------------------------------------------------
 */ 


#include "stdafx.h"
#include "ItemQuery.h"
#include "ItemQueryDlg.h"
#include "ItemViewDlg.h"

#import "qbFC13.dll" no_namespace, named_guids
CString QBFCLatestVersion(IQBSessionManagerPtr SessionManager);
IMsgSetRequestPtr GetLatestMsgSetRequest(IQBSessionManagerPtr SessionManager);

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CItemQueryDlg dialog

CItemQueryDlg::CItemQueryDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CItemQueryDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CItemQueryDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CItemQueryDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CItemQueryDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CItemQueryDlg, CDialog)
	//{{AFX_MSG_MAP(CItemQueryDlg)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CItemQueryDlg message handlers

BOOL CItemQueryDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

  return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CItemQueryDlg::OnPaint() 
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

HCURSOR CItemQueryDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

#define APPID        "123"
#define APPNAME      "IDN QBFC VC ItemQuery"

void CItemQueryDlg::OnOK() 
{
  try {
	 
    //Create session manager
    IQBSessionManagerPtr sessionManager(CLSID_QBSessionManager);
    sessionManager->OpenConnection ( APPID, APPNAME );
    sessionManager->BeginSession ( "", omDontCare );

    //Create message request set object
    IMsgSetRequestPtr requestMsgSet = GetLatestMsgSetRequest(sessionManager);
    requestMsgSet->Attributes->OnError = roeContinue;
    
    //Append the request to the request set object
    IItemQueryPtr itemQ = requestMsgSet->AppendItemQueryRq();
    itemQ->ORListQuery->ListFilter->MaxReturned->SetValue( 30 );
    
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
  
     
    //The detail property of IResponse object is a list for all queries,
    //except CompanyQuery, HostQuery, PreferencesQuery, which don't return a list,
    //but a single Ret object.
    //In this case the list in detail is an IORItemRetList (a list of IORItemRet objects)
    
    //For help finding out the detail's type, use the following line:
    //_bstr_t detailType =  response->Detail->Type->GetAsString();
    
    IORItemRetListPtr orItemRetList = response->Detail;
    if ( orItemRetList != NULL ) 
    {
      int ndxMax = orItemRetList->Count ;
        
      for( int ndx = 0; ndx < ndxMax ; ++ndx )
      {
        IORItemRetPtr orItemRet = orItemRetList->GetAt(ndx);
        ENORItemRet orType = orItemRet->ortype;
          
        //The ortype property returns an enum
        //of the elements that can be contained in the OR object
        switch ( orType )
        {
        case orirItemServiceRet: //"orir" prefix comes from OR + Item + Ret
          {
            IItemServiceRetPtr itemServiceRet = orItemRet->ItemServiceRet;
            msg += "\r\nItemServiceRet: FullName = " ;
            msg += itemServiceRet->FullName->GetValue();

                
            //Retrieving a field that may not be set
            //in this case a reference object (IQBBaseRef)
            if ( itemServiceRet->SalesTaxCodeRef != NULL ) 
            {
              msg += ", SalesTaxCodeRef = " ;
              msg += itemServiceRet->SalesTaxCodeRef->FullName->GetValue();
            }
          }
          break;
        case orirItemInventoryRet:
          {
            IItemInventoryRetPtr itemInventoryRet = orItemRet->ItemInventoryRet;
            msg += "\r\nItemInventoryRet: FullName = ";
            msg += itemInventoryRet->FullName->GetValue();
                
            //Retrieving a field that may not be set,
            //in this case, a quantity
            if ( itemInventoryRet->QuantityOnHand != NULL )
            {
              msg += ", QuantityOnHand = ";
              msg += itemInventoryRet->QuantityOnHand->GetAsString();
            }
          }
          break;
        case orirItemNonInventoryRet:
          {      
            IItemNonInventoryRetPtr itemNonInventoryRet = orItemRet->ItemNonInventoryRet;
            msg += "\r\nItemNonInventoryRet: FullName = " ;
            msg += itemNonInventoryRet->FullName->GetValue();
          }
          break;
        default:
          {
            //...
            // Have specific code for the other item types.
            msg += "\r\nOther Item Type";
          }
        }      
      }
    }
    
    sessionManager->EndSession();
    sessionManager->CloseConnection();
    
    //Show item list
    ItemViewDlg dlg(msg );
    dlg.DoModal();

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
