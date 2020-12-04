
/* InvoiceQueryDlg.cpp : implementation file
 *
 * Description:  This sample demonstrates the use of QBFC, by querying 
 *               for invoices within a given date range.
 *               It includes examples of the following:
 *                  - Constructing a complex query
 *                  - Looping through the list of invoices in a response
 *                  - Getting detailed invoice information (invoice lines)
 *                  - Reading data from an OR object
 *                  - Checking fields that are not guaranteed in the response
 *                    and obtaining data from them
 *
 * Created On: 08/15/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 *     8/9/2013 updated to qbfc13
*/

#include "stdafx.h"
#include "InvoiceQuery.h"
#include "InvoiceQueryDlg.h"
#include "InvoiceViewDlg.h"

#import "qbFC13.dll" no_namespace, named_guids
CString QBFCLatestVersion(IQBSessionManagerPtr SessionManager);
IMsgSetRequestPtr GetLatestMsgSetRequest(IQBSessionManagerPtr SessionManager);


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CInvoiceQueryDlg dialog

CInvoiceQueryDlg::CInvoiceQueryDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CInvoiceQueryDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CInvoiceQueryDlg)
	m_includeLineItems = TRUE;
	m_fromTxnDate = _T("");
	m_toTxnDate = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CInvoiceQueryDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CInvoiceQueryDlg)
	DDX_Check(pDX, IDC_IncludeLineItems, m_includeLineItems);
	DDX_Text(pDX, IDC_FromTxnDate, m_fromTxnDate);
	DDX_Text(pDX, IDC_ToTxnDate, m_toTxnDate);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CInvoiceQueryDlg, CDialog)
	//{{AFX_MSG_MAP(CInvoiceQueryDlg)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CInvoiceQueryDlg message handlers

BOOL CInvoiceQueryDlg::OnInitDialog()
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

void CInvoiceQueryDlg::OnPaint() 
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
HCURSOR CInvoiceQueryDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CInvoiceQueryDlg::OnOK() 
{
  UpdateData(TRUE); //transfer data from controls to our variables

  COleDateTime oleFromTxnDate ;
  COleDateTime oleToTxnDate;

  //Verify format of the dates
  if ( !m_fromTxnDate.IsEmpty() && 
       !oleFromTxnDate.ParseDateTime( m_fromTxnDate, VAR_DATEVALUEONLY ) )
  {
    AfxMessageBox( _T("From Txn Date format is wrong!") );
    return ;
  } 
  
  if ( !m_toTxnDate.IsEmpty() && 
       !oleToTxnDate.ParseDateTime( m_toTxnDate, VAR_DATEVALUEONLY ) )
  {
    AfxMessageBox( _T("To Txn Date format is wrong!") );
    return ;
  } 
  
  QueryInvoices( !m_fromTxnDate.IsEmpty() ? &oleFromTxnDate : NULL, 
    !m_toTxnDate.IsEmpty() ? &oleToTxnDate : NULL, m_includeLineItems );

}

#define APPID "123"
#define APPNAME "IDN QBFC VC++ InvoiceQuery"

#define QBXML_MAJORVER 2
#define QBXML_MINORVER 0

CString GetInvoiceRetDetail(IInvoiceRetPtr& invRet);

void CInvoiceQueryDlg::QueryInvoices( COleDateTime* oleFromTxnDate, 
    COleDateTime* oleToTxnDate, BOOL m_includeLineItems )
{
   
  try {
    //Step 1: Start session with QuickBooks
    IQBSessionManagerPtr sessionManager(CLSID_QBSessionManager);
    sessionManager->OpenConnection ( APPID, APPNAME );
    sessionManager->BeginSession ( "", omDontCare );

    //Step 2: Create Message Set request
    IMsgSetRequestPtr requestMsgSet = GetLatestMsgSetRequest(sessionManager);
    requestMsgSet->Attributes->OnError = roeContinue;
    
    //Step 3: Create the query object needed to perform InvoiceQueryRq
    IInvoiceQueryPtr invQuery = requestMsgSet->AppendInvoiceQueryRq();
    
    IInvoiceFilterPtr invFilter = invQuery->ORInvoiceQuery->InvoiceFilter;
    
    //If there's date info, create the filter and put it in the query
    if( oleFromTxnDate != NULL )
    {
      invFilter->ORDateRangeFilter->TxnDateRangeFilter->
        ORTxnDateRangeFilter->TxnDateFilter->
        FromTxnDate->SetValue (*oleFromTxnDate);
        
    }
    if( oleToTxnDate != NULL) {
      invFilter->ORDateRangeFilter->TxnDateRangeFilter->
        ORTxnDateRangeFilter->TxnDateFilter->
        ToTxnDate->SetValue (*oleToTxnDate);
    }
    invQuery->IncludeLineItems->SetValue (
      m_includeLineItems ? VARIANT_TRUE: VARIANT_FALSE);
    
        
    //Step 4: Do the request
    IMsgSetResponsePtr responseMsgSet = sessionManager->DoRequests(requestMsgSet);
    
    //Close connection because we don't need it anymore
    sessionManager->EndSession();
    sessionManager->CloseConnection();
    
    // Uncomment the following to see the request and response XML for debugging
    //_bstr_t rq = requestMsgSet->ToXMLString();
    //_bstr_t rs = responseMsgSet->ToXMLString();
    
    //Step 5: Interpret the response
    IResponseListPtr rsList = responseMsgSet->ResponseList;
    
    //Retrieve the one response corresponding to our single request
    IResponsePtr response = rsList->GetAt(0);

    if (response->StatusCode != 0) 
    {
      if (response->StatusCode == 1)  //No record found
      {
        AfxMessageBox ("No invoice is found");
      }
      else
      {
        CString msg;
        msg.Format ( "Error occured.  Status Code = %d, Status Message = %S"
          ", Status Severity = %S", response->StatusCode, 
          (BSTR)response->StatusMessage, (BSTR)response->StatusSeverity );
        AfxMessageBox (msg );
      }   
      return;
    } 
    else //We have one or more invoices-> show them
    {
      IInvoiceRetListPtr invoiceList = response->Detail;
      int maxCnt =  invoiceList->Count; 

      CString msg;
      const int cMaxShown = 30;
    
      if ( maxCnt > cMaxShown )
      {
          msg.Format("Showing %d out of %d invoices\r\n", cMaxShown, maxCnt);
      }

      for( int ndx = 0;  ndx < maxCnt; ndx ++ )
      {
        if( ndx == cMaxShown ) {
          break;
        }
        IInvoiceRetPtr invoiceRet = invoiceList->GetAt(ndx);
        //Add to message
        msg += GetInvoiceRetDetail( invoiceRet );
      }

      InvoiceViewDlg dlg(msg );
      dlg.DoModal();

    }
  } 
  catch ( _com_error e )
  {
    CString msg;
    _bstr_t desc = e.Description();
    msg.Format( "Error: 0x%x, %s" , e.Error(), 
      (desc.length() > 0) ? (const char*)desc : "" );
    AfxMessageBox( msg );
  }

}

CString GetInvoiceRetDetail(IInvoiceRetPtr& invRet) {

  CString msg;
    
  //Retrieve guaranteed fields
  _bstr_t txnNumber = invRet->TxnNumber->GetValue();
  _bstr_t custFullName = invRet->CustomerRef->FullName->GetValue() ;
  msg.Format( "TxnNumber = %S, Customer = %S", 
    (BSTR)txnNumber, (BSTR)custFullName );
    
  //Retrive non-guaranteed fields
  if ( invRet->RefNumber != NULL ) 
  {
    msg += ", RefNumber = " ;
    msg += (const char*)invRet->RefNumber->GetValue();
  }
    
  if( invRet->Memo  != NULL )
  {
    msg += ", Memo = ";
    msg += (const char*)invRet->Memo->GetValue();
  }
    
   
  //Retrieve invoice line list
  //Each line can be either InvoiceLineRet OR InvoiceLineGroupRet
  IORInvoiceLineRetListPtr orInvoiceLineRetList = invRet->ORInvoiceLineRetList;
  if (orInvoiceLineRetList != NULL )
  {
    
    int linendxMax = orInvoiceLineRetList->Count;
    for( int linendx = 0; linendx < linendxMax; ++ linendx )
    {
      IORInvoiceLineRetPtr orInvoiceLineRet = 
        orInvoiceLineRetList->GetAt(linendx);
            
      CString tmp;
      tmp.Format ( "\r\n\t Line: %d",  linendx );
      msg += tmp;

      //Check what to retrieve from the orInvoiceLineRet object
      //based on the "ortype" property
      
      if (orInvoiceLineRet->ortype == orilrInvoiceLineRet) 
      {
             
        msg += ", TxnLineId: " ;
        msg += (const char *)orInvoiceLineRet->InvoiceLineRet->TxnLineID->GetValue();

        if (orInvoiceLineRet->InvoiceLineRet->Desc  != NULL )
        {
          msg += ", Desc: " ;
          msg += (const char *)orInvoiceLineRet->InvoiceLineRet->Desc->GetValue();
        }
        if (orInvoiceLineRet->InvoiceLineRet->Amount != NULL) 
        {
          msg += ", Amount: " ;
          msg += (const char *)orInvoiceLineRet->InvoiceLineRet->Amount->GetAsString();
        }
                                
        if (orInvoiceLineRet->InvoiceLineRet->ItemRef != NULL)
        {
           msg += ", Quantity: " ;
           msg += (const char *)orInvoiceLineRet->InvoiceLineRet->ItemRef->FullName->GetValue();
        }
                
        IORRatePtr orRate = orInvoiceLineRet->InvoiceLineRet->ORRate;

        if (orRate != NULL) 
        {
          if ( orRate->ortype == orrRate )
          {
            msg += ", Rate: " ;
            msg += (const char *)orRate->Rate->GetAsString();
          }
          else
          {
            msg += ", RatePercent: " ;
            msg += (const char *)orRate->RatePercent->GetAsString();
          }
        } 
      }
      else // orInvoiceLineRet.ortype == orilrInvoiceLineGroupRet
      {
        msg += ", Group Name: " ;
        msg += (const char *)orInvoiceLineRet->InvoiceLineGroupRet->ItemGroupRef->FullName->GetValue();

        msg += ", Total Amount: " ;
        msg += (const char *)orInvoiceLineRet->InvoiceLineGroupRet->TotalAmount->GetAsString();
         
        if (orInvoiceLineRet->InvoiceLineGroupRet->Desc != NULL ) 
        {
          msg += ", Desc: " ;
          msg += (const char *)orInvoiceLineRet->InvoiceLineGroupRet->Desc->GetValue();
        }
      } //ottype         
    } //for     
  }//  orInvoiceLineRetList != NULL 
    
  msg += "\r\n" ;
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
