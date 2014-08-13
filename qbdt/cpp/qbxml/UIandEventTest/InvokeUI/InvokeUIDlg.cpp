/*---------------------------------------------------------------------------
 * FILE: InvokeUIDlg.cpp
 *
 * Description:
 * MFC-generated dialog class.
 *
 * Created On: 09/15/2003
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#include "stdafx.h"
#include "InvokeUI.h"
#include "InvokeUIDlg.h"
#include "qbXMLTags.h"
#include "Utility.h"
#include "qbXMLObject.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

//-------------------------------------------------------------------
// CInvokeUIDlg dialog

// MFC-added contructor.
CInvokeUIDlg::CInvokeUIDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CInvokeUIDlg::IDD, pParent)
{
    // MFC-managed elements.
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	//{{AFX_DATA_INIT(CInvokeUIDlg)
	//}}AFX_DATA_INIT

    // Lists to hold transaction and item objects displayed
    // in Combo boxes. ComboBox members will point to these
    // objects.
    m_pTxnList = new qbXMLObjectList();;
    m_pItemList = new qbXMLObjectList();
}

// Destructor.
CInvokeUIDlg::~CInvokeUIDlg()
{
    // Delete the lists. These objects take care of 
    // deleting member objects.
    if( m_pTxnList ) {
        delete m_pTxnList;
        m_pTxnList = NULL;
    }
    if( m_pItemList ) {
        delete m_pItemList;
        m_pItemList = NULL;
    }
}


// MFC-added data exchange
void CInvokeUIDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CInvokeUIDlg)
	//}}AFX_DATA_MAP
}

// MFC-managed message map.
BEGIN_MESSAGE_MAP(CInvokeUIDlg, CDialog)
	//{{AFX_MSG_MAP(CInvokeUIDlg)
	ON_BN_CLICKED(IDC_TXN_TYPE_VIEW, OnTxnTypeView)
	ON_BN_CLICKED(IDC_TXN_TYPE_LOAD, OnTxnTypeLoad)
	ON_BN_CLICKED(IDC_TXN_LIST_VIEW, OnTxnListView)
	ON_BN_CLICKED(IDC_ITEM_TYPE_VIEW, OnItemTypeView)
	ON_BN_CLICKED(IDC_ITEM_TYPE_LOAD, OnItemTypeLoad)
	ON_BN_CLICKED(IDC_ITEM_LIST_VIEW, OnItemListView)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


//-------------------------------------------------------------------
// CInvokeUIDlg message handlers

// OnInitDialog()
// Dialog initialization.
BOOL CInvokeUIDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	return TRUE;  // return TRUE  unless you set the focus to a control
}

// Handlers for dialog buttons.
// Transactions and Items use the same helper methods
// by passing to those methods the control ID's and 
// item XML tags for the appropriate type.
void CInvokeUIDlg::OnTxnTypeView() 
{
    OpenAddWindow( IDC_TXN_TYPE, "TxnDisplayAdd" );
}
void CInvokeUIDlg::OnTxnTypeLoad() 
{
    LoadList( IDC_TXN_TYPE, IDC_TXN_LIST, m_pTxnList );
}
void CInvokeUIDlg::OnTxnListView() 
{
    OpenModWindow( IDC_TXN_LIST, "TxnDisplayMod", "TxnID" );
}
void CInvokeUIDlg::OnItemTypeView() 
{
    OpenAddWindow( IDC_ITEM_TYPE, "ListDisplayAdd" );
}
void CInvokeUIDlg::OnItemTypeLoad() 
{
    LoadList( IDC_ITEM_TYPE, IDC_ITEM_LIST, m_pItemList );
}
void CInvokeUIDlg::OnItemListView() 
{
    OpenModWindow( IDC_ITEM_LIST, "ListDisplayMod", "ListID" );
}

//-------------------------------------------------------------------
// Helper methods.

// LoadList()
// Load item list with data from QuickBooks.
//  nTypeList = id of control to obtain type description from
//  nItemList = id of control to add QB items to
//  pList     = member variable that holds QB item objects
//
void CInvokeUIDlg::LoadList( UINT nTypeList, 
                             UINT nItemList, 
                             qbXMLObjectList* pList ) {
    // Get list that we'll be adding items to.
    CComboBox* pCombo = (CComboBox*)GetDlgItem( nItemList );
    if( pCombo == NULL ) {
        return;
    }

    // Get currently selected Transaction or ListItem type.
    CString strType;
    int nLength = GetDlgItemText( nTypeList, strType );
    if( nLength == 0 ) {
        AfxMessageBox( "No type selected." );
        return;
    }

    // Query for data. 
    // If query fails, bail out (error already displayed).
    string strResponse;
    if( !DoQueryRequest(strType,strResponse) ) { 
        return;
    }

    // Load objects from response XML.
    pCombo->ResetContent();
    pList->removeAll();
    pList->parseObjects( strResponse, strType );
    int count = pList->getSize();
    if( count > 0 ) {
        for( int i = 0; i < count; i++ ) {
            qbXMLObject* pObject = pList->getItemAt( i );
            string strItem = strType;
            strItem += ": ";
            strItem += pObject->getDisplayName();
            int iItem = pCombo->AddString( strItem.c_str() );
            // Save pointer to object so that we can get
            // it later from a selected item.
            pCombo->SetItemDataPtr( iItem, pObject );
        }
    }
    // If no objects obtained, display message.
    else {
        CString strMessage = "No \"";
        strMessage += strType;
        strMessage += " \" objects found.";
        AfxMessageBox( strMessage );
    }

    // Select first item in list.
    pCombo->SetCurSel( 0 );

}

// OpenAddWindow()
// Open a new form in QB (using TxnDisplayAdd or ListDisplayAdd).
//   nTypeList = control id to read type of item from
//   lpszTag   = XML query tag for this type of item
//
void CInvokeUIDlg::OpenAddWindow( UINT nTypeList, LPCTSTR lpszTag ) {
    // Get currently selected Transaction or ListItem type.
    CString strType;
    int nLength = GetDlgItemText( nTypeList, strType );
    if( nLength == 0 ) {
        AfxMessageBox( "No type selected." );
        return;
    }

    // Build and send request to invoke selected form.
    DoInvokeRequest( lpszTag, strType, NULL, NULL );
}

// OpenModWindow()
// Open an existing form in QB (using TxnDisplayMod or ListDisplayMod).
//   nList     = control id to read item info from
//   lpszTag   = XML query tag for this type of item
//   lpszIDTag = XML tag for ID element of query
//
void CInvokeUIDlg::OpenModWindow( UINT nList, 
                                  LPCTSTR lpszTag, 
                                  LPCTSTR lpszIDTag ) {
    // Get currently selected Transaction or ListItem type.
    CString strType;
    int nLength = GetDlgItemText( nList, strType );
    if( nLength == 0 ) {
        AfxMessageBox( "No item selected." );
        return;
    }

    // Get object associated with currently selected item.
    CComboBox* pCombo = (CComboBox*)GetDlgItem( nList );
    int nSel = pCombo->GetCurSel();
    qbXMLObject* pObject = (qbXMLObject*)pCombo->GetItemDataPtr( nSel );
    if( pObject ) {
        // Build and send request to invoke selected form.
        DoInvokeRequest( lpszTag, 
                         pObject->getType().c_str(), 
                         lpszIDTag, 
                         pObject->getID().c_str() );
    }
}

// DoInvokeRequest()
// Build and send request to invoke a new or existing form.
//   lpszTag     = main query tag for this item type
//   lpszType    = type of object being invoked
//   lpszIDTag   = query tag for ID element
//                 ( NULL for invoking a new form )
//   lpszIDValue = value for ID element
//                 ( NULL for invoking a new form )
// Returns true if request succeeded; false if not.
//
BOOL CInvokeUIDlg::DoInvokeRequest( LPCTSTR lpszTag,
                                    LPCTSTR lpszType,
                                    LPCTSTR lpszIDTag,
                                    LPCTSTR lpszIDValue ) {
	CWaitCursor waitCursor;

	// See if user wants to display response and request XML.
	BOOL showXML = IsDlgButtonChecked( IDC_SHOW_XML );

    // Make sure company file is opened.
    qbXMLRPWrapper  requestProcessor;
    requestProcessor.OpenCompanyFile( "" );
    // If an error occurred, display error message.
	if( requestProcessor.GetHasError() ) {
        DisplayRequestError( requestProcessor, 
                             "Unable to connect to open data file." );
        return false;
    }


	// Build "query" request.
	string strRequest = BuildInvokeRequest( lpszTag, 
                                            lpszType, 
                                            lpszIDTag, 
                                            lpszIDValue );
	// If an error occurred building request, bail out (error was already displayed).
	if( strRequest.length() == 0 ) {
		return false;
	}
    // Display request XML if requested.
    else if( showXML ) {
        Utility::ShowXML( "Invoke UI Request",
                          "Request to invoke UI within QuickBooks",
                          strRequest.c_str() );
    }

    waitCursor.Restore();

	// Send "query" request.
	string strResponse = requestProcessor.ProcessRequest( strRequest );
	// If an error occurred with request, display error and bail out.
	if( requestProcessor.GetHasError() ) {
        DisplayRequestError( requestProcessor, 
                             "Error sending Invoke UI request to QuickBooks." );
		return false;
	}
    // Display response XML if requested.
    else if( showXML ) {
        Utility::ShowXML( "Invoke UI Response",
                          "Response from invoking UI within QuickBooks",
                          strResponse.c_str() );
    }

    // If the response contains an error, bail out (message already displayed).
    if( Utility::DoesResponseHaveError(QBXMLMSGSRS_TAG,strResponse) ) {
        return false;
    }

	return true;
}

// BuildInvokeRequest()
// Builds XML and checks for errors for DoInvokeRequest().
//   lpszTag     = main query tag for this item type
//   lpszType    = type of object being invoked
//   lpszIDTag   = query tag for ID element
//                 ( NULL for invoking a new form )
//   lpszIDValue = value for ID element
//                 ( NULL for invoking a new form )
// Returns request XML if it could be created.
//
string CInvokeUIDlg::BuildInvokeRequest( LPCTSTR lpszTag,
                                         LPCTSTR lpszType,
                                         LPCTSTR lpszIDTag,
                                         LPCTSTR lpszIDValue ) {

    // Get tags for this type of query.
    string strRqTag = lpszTag;    // Request tag.
    strRqTag += "Rq";
    string strTypeTag = lpszTag;  // Form/List type tag.
    strTypeTag += "Type";


	// Begin request.
	DOMXMLBuilder xmlBuilder;
	xmlBuilder.CreateXMLHeader();

	// Check initial xmlBuilder access for COM and other initialization errors.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

    // Start request aggregate.
	xmlBuilder.AddAggregate( strRqTag );

    // Add type element.
    xmlBuilder.AddElement( strTypeTag, lpszType );

    // If this is a mod request, add ID tag.
    if( lpszIDTag && lpszIDValue ) {
        xmlBuilder.AddElement( lpszIDTag, lpszIDValue );
    }

    // Finish request aggregate.
	xmlBuilder.AddEndAggregate( strRqTag );
    
	// Finish request.
	xmlBuilder.CreateXMLTrailer();

	// If any errors occurred, display error message.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

	// If no errors, return request formatted with tabs and line breaks.
	return xmlBuilder.GetXML();
}

// DoQueryRequest()
// Build and send request to get data from QuickBooks.
//   lpszType    = type of object being queried
//   strResponse = return value for generated response
// Returns TRUE if request was successful; FALSE if any errors occurred
//
BOOL CInvokeUIDlg::DoQueryRequest( LPCTSTR lpszType, string& strResponse ) {
	CWaitCursor waitCursor;

	// See if user wants to display response and request XML.
	BOOL showXML = IsDlgButtonChecked( IDC_SHOW_XML );

    // Make sure company file is opened.
    qbXMLRPWrapper requestProcessor;
    requestProcessor.OpenCompanyFile( "" );
    // If an error occurred, display error message.
	if( requestProcessor.GetHasError() ) {
        DisplayRequestError( requestProcessor,
                             "Unable to connect to open data file." );
        return false;
    }


	// Build "query" request.
	string strRequest = BuildQueryRequest( lpszType );
	// If an error occurred building request, bail out (error was already displayed).
	if( strRequest.length() == 0 ) {
		return false;
	}
    // Display request XML if requested.
    else if( showXML ) {
        Utility::ShowXML( "Query Data Request",
                          "Request to retrieve data from QuickBooks",
                          strRequest.c_str() );
    }

    waitCursor.Restore();

	// Send "query" request.
	strResponse = requestProcessor.ProcessRequest( strRequest );
	// If an error occurred with request, display error and bail out.
	if( requestProcessor.GetHasError() ) {
        DisplayRequestError( requestProcessor,
                             "Error requesting data from QuickBooks." );
		return false;
	}
    // Display response XML if requested.
    else if( showXML ) {
        Utility::ShowXML( "Query Data Response",
                          "Response for retrieving data from QuickBooks",
                          strResponse.c_str() );
    }

    // If the response contains an error, bail out (message already displayed).
    if( Utility::DoesResponseHaveError(QBXMLMSGSRS_TAG,strResponse) ) {
        return false;
    }

	return true;
}

// BuildQueryRequest()
// Builds XML and checks for errors for DoQueryRequest().
//   lpszType    = type of object being queried
// Returns request XML if it could be created.
//
string CInvokeUIDlg::BuildQueryRequest( LPCTSTR lpszType ) {

    // Get tag for this type of query.
    string strRqTag = lpszType;
    strRqTag += "QueryRq";

	// Begin request.
	DOMXMLBuilder xmlBuilder;
	xmlBuilder.CreateXMLHeader();

	// Check initial xmlBuilder access for COM and other initialization errors.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

    // Create RequestID attribute.
    string attrName[] = { "requestID" };
    string attrValue[] = { "123" };

    // Start request aggregate.
	xmlBuilder.AddAggregate( strRqTag, attrName, attrValue, 1 );

    // Finish request aggregate.
	xmlBuilder.AddEndAggregate( strRqTag );
    
	// Finish request.
	xmlBuilder.CreateXMLTrailer();

	// If any errors occurred, display error message.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

	// If no errors, return request formatted with tabs and line breaks.
	return xmlBuilder.GetXML();
}

// DisplayRequestError()
// Displays an error message based on a request error.
// Gets error message from class member request processor object.
void CInvokeUIDlg::DisplayRequestError( qbXMLRPWrapper& requestProcessor,
                                        LPCTSTR lpszMessage ) {
	string errMsg = lpszMessage;
    if( requestProcessor.GetHasError() ) {
	    errMsg += "\n\n";
	    errMsg += requestProcessor.GetErrorMsg().c_str();
	    errMsg += "\n\nThere may be additional info in the qbsdklog.txt file (in the QuickBooks executable directory).";
    }
	AfxMessageBox( errMsg.c_str() );
}

// DisplayXMLError()
// Displays an error message based on an XML DOM error.
void CInvokeUIDlg::DisplayXMLError( DOMXMLBuilder& xmlBuilder ) {
	string errMsg( "Error building XML.\n\n" );
	errMsg.append( xmlBuilder.GetErrorMsg() );
	AfxMessageBox( errMsg.c_str() );
}
