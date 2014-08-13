/*---------------------------------------------------------------------------
 * FILE: EventSubscribeDlg.cpp
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
#include "EventSubscribe.h"
#include "EventSubscribeDlg.h"
#include "DOMXMLBuilder.h"
#include "qbXMLTags.h"
#include "qbXMLRPWrapper.h"
#include "Utility.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

//-------------------------------------------------------------------
// Constants
static string strListItemDescription = "(list item)";
static string strTransactionDescription = "(transaction)";

//-------------------------------------------------------------------
// CEventSubscribeDlg dialog

// MFC-add contructor.
CEventSubscribeDlg::CEventSubscribeDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CEventSubscribeDlg::IDD, pParent)
{
    // MFC-managed elements.
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	//{{AFX_DATA_INIT(CEventSubscribeDlg)
	//}}AFX_DATA_INIT
}

// MFC-added data exchange
void CEventSubscribeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CEventSubscribeDlg)
	//}}AFX_DATA_MAP
}

// MFC-managed message map.
BEGIN_MESSAGE_MAP(CEventSubscribeDlg, CDialog)
	//{{AFX_MSG_MAP(CEventSubscribeDlg)
	ON_BN_CLICKED(IDC_ADD_SUBSCRIPTION, OnAddSubscription)
	ON_BN_CLICKED(IDC_TXN_SUBSCRIBE, OnTxnSubscribe)
	ON_BN_CLICKED(IDC_LIST_SUBSCRIBE, OnListSubscribe)
	ON_BN_CLICKED(IDC_REMOVE, OnRemove)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


//-------------------------------------------------------------------
// CEventSubscribeDlg message handlers

// OnInitDialog()
// Dialog initialization.
BOOL CEventSubscribeDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

    CheckDlgButton( IDC_LAUNCH_HANDLER, TRUE );
    

    return TRUE;  // return TRUE  unless you set the focus to a control
}

// OnTxnSubscribe()
// Add a transaction event subscription to list.
void CEventSubscribeDlg::OnTxnSubscribe() 
{
    AddEvent( strTransactionDescription, IDC_TXN, IDC_TXN_ADD, IDC_TXN_MOD, IDC_TXN_DEL, 0 );
}

// OnListSubscribe()
// Add a list item event subscription to list.
void CEventSubscribeDlg::OnListSubscribe() 
{
    AddEvent( strListItemDescription, IDC_LIST, IDC_LIST_ADD, IDC_LIST_MOD, IDC_LIST_DEL, IDC_LIST_MERGE );
}

// OnRemove()
// Remove currently selected event subscription from list.
void CEventSubscribeDlg::OnRemove() 
{
	CListBox* pList = (CListBox*)GetDlgItem( IDC_SELECTED );
    int curSel = pList->GetCurSel();
    if( curSel == LB_ERR ) {
        AfxMessageBox( "No item is selected." );
    }
    else {
        pList->DeleteString( curSel );
    }
}

// OnAddSubscription()
// Handle button click to add subscription.
void CEventSubscribeDlg::OnAddSubscription()
{
	// If no subscriptions are defined, display error and bail out.
    if( !HasDataEventSubscriptions() && !HasUIEventSubscriptions() ) {
		AfxMessageBox( "No subscriptions are defined." );
		return;
	}

	qbXMLRPWrapper requestProcessor;
	BOOL alreadyExists = false;

	// There is no request to directly replace an existing subscription.
	// So first see if a subscription already exists.
	BOOL success = DoQueryRequest( requestProcessor, alreadyExists );

	// If a subscription does already exist, delete it.
	if( success && alreadyExists ) {
		success = DoDeleteRequest( requestProcessor );
	}

	// Now add the subscription.
	if( success ) {
		success = DoAddRequest( requestProcessor );
	}

    // Remind user to restart QuickBooks.
	if( success ) {
		AfxMessageBox( "Subscription added.\nClose and re-open any opened data file to enable new event(s)." );
	}
}

//-------------------------------------------------------------------
// Helper methods.

// DoQueryRequest()
// Build and send request to query for a menu subscription.
//   requestProcessor = QBXML request processor used to perform requests
//   alreadyExists    = used to return whether subscription already exists
// Returns true if request succeeded; false if not.
//
BOOL CEventSubscribeDlg::DoQueryRequest( qbXMLRPWrapper& requestProcessor, 
							             BOOL& alreadyExists ) {
	// See if user wants to display response and request XML.
	BOOL showXML = IsDlgButtonChecked( IDC_SHOW_XML );

	// Initialize current "exists" status.
	alreadyExists = false;

	// Build "query subscription" request.
	string strRequest = BuildQueryRequest();
	// If an error occurred building request, bail out (error was already displayed).
	if( strRequest.length() == 0 ) {
		return false;
	}
	// Show request.
	else if( showXML ) {
        Utility::ShowXML( "Event \"query subscription\" request",
                          "Request to check if event subscriptions already exist.",
                          strRequest.c_str() );
	}

	CWaitCursor waitCursor;

	// Send "query subscription" request.
	string strResponse = requestProcessor.ProcessSubscription( strRequest );
	// If an error occurred with request, display error and bail out.
	if( requestProcessor.GetHasError() ) {
		DisplayRequestError( "Error checking for existing subscription", requestProcessor );
		return false;
	}
		// Show response.
	else if( showXML ) {
        Utility::ShowXML( "Event \"query subscription\" response",
                          "Response from checking for existing subscriptions.",
                          strResponse.c_str() );
	}

    // If the response contains an error, display it and bail out.
    if( Utility::DoesResponseHaveError(QBXMLSUBMSGSRS_TAG,strResponse) ) {
        return false;
    }

    // See if subscriptions we're creating already exist.
    // We check for an existing subscription by searching for the
    // ???EventSubscriptionRet tag, which is returned only if a
    // matching subscription already exists.
    BOOL overwriteUIEvent = false;
    if( HasUIEventSubscriptions() ) {
        overwriteUIEvent = ( (int)strResponse.find("UIEventSubscriptionRet") > 0 );
    }
    BOOL overwriteDataEvent = false;
    if( HasDataEventSubscriptions() ) {
        overwriteDataEvent = ( (int)strResponse.find("DataEventSubscriptionRet") > 0 );
    }

	// If subscriptions already exists, prompt to overwrite (since there can
	// only be one subscription of each subscription type per "subscriber id".)
	if( overwriteDataEvent || overwriteUIEvent ) {
		// Set current status;
		alreadyExists = true;
		// Display prompt.
		string strMessage = "Event subscriptions already exist.\n";
		strMessage += "Do you wish to overwrite ?";
		// If user does not want to overwrite, bail out.
		if( IDNO == AfxMessageBox(strMessage.c_str(),MB_YESNO) ) {
			return false;
		}
	}

	return true;
}

// DoDeleteRequest()
// Build and send request to delete a subscription.
//   requestProcessor = QBXML request processor used to perform requests
// Returns true if request succeeded; false if not.
//
BOOL CEventSubscribeDlg::DoDeleteRequest( qbXMLRPWrapper& requestProcessor ) {
	// See if user wants to display response and request XML.
	BOOL showXML = IsDlgButtonChecked( IDC_SHOW_XML );

	// Build subscription delete request.
	string strRequest = BuildDeleteRequest();
	// If an error occurred building request, bail out (error was already displayed).
	if( strRequest.length() == 0 ) {
		return false;
	}
	// Show request.
	else if( showXML ) {
        Utility::ShowXML( "Event \"delete subscription\" request",
                          "Request to delete existing event subscriptions.",
                          strRequest.c_str() );
	}

	CWaitCursor waitCursor;

	// Send "delete subscription" request.
	string strResponse = requestProcessor.ProcessSubscription( strRequest );
	// If an error occurred with request, display error and bail out.
	if( requestProcessor.GetHasError() ) {
		DisplayRequestError( "Error deleting existing subscription", requestProcessor );
		return false;
	}
	// If no error occurred, display response.
	else if( showXML ) {
        Utility::ShowXML( "Event \"delete subscription\" response",
                          "Response from deleting event subscriptions.",
                          strResponse.c_str() );
	}

    // If the response contains an error, display it and bail out.
    if( Utility::DoesResponseHaveError(QBXMLSUBMSGSRS_TAG,strResponse) ) {
        return false;
    }

	return true;
}

// DoAddRequest()
// Build and send request to add a subscription.
//   requestProcessor = QBXML request processor used to perform requests
// Returns true if request succeeded; false if not.
//
BOOL CEventSubscribeDlg::DoAddRequest( qbXMLRPWrapper& requestProcessor ) {
	// See if user wants to display response and request XML.
	BOOL showXML = IsDlgButtonChecked( IDC_SHOW_XML );

	// Build "add subscription" request.
	string strRequest = BuildAddRequest();
	// If an error occurred building request, bail out (error was already displayed).
	if( strRequest.length() == 0 ) {
		return false;
	}
	// Show request.
	else if( showXML ) {
        Utility::ShowXML( "Event \"add subscription\" request",
                          "Request to add event subscription(s).",
                          strRequest.c_str() );
	}

	CWaitCursor waitCursor;

    // Send "add subscription" request.
	string strResponse = requestProcessor.ProcessSubscription( strRequest );
	// If an error occurred with request, display error and bail out.
	if( requestProcessor.GetHasError() ) {
		DisplayRequestError( "Error adding subscription", requestProcessor );
		return false;
	}
	// If no error occurred, display response.
	else if( showXML ) {
        Utility::ShowXML( "Event \"add subscription\" response",
                          "Response from adding event subscription(s).",
                          strResponse.c_str() );
	}

    // If the response contains an error, display it and bail out.
    if( Utility::DoesResponseHaveError(QBXMLSUBMSGSRS_TAG,strResponse) ) {
        return false;
    }

	return true;
}

// BuildQueryRequest()
// Builds XML and checks for errors for DoQueryRequest().
// Returns request XML if it could be created.
//
string CEventSubscribeDlg::BuildQueryRequest() {
	// Begin request.
	DOMXMLBuilder xmlBuilder;
	xmlBuilder.CreateSubscriptionXMLHeader();

	// Check initial xmlBuilder access for COM and other initialization errors.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

    // If there are UI event subscriptions,
    // add UI event query request.
    if( HasUIEventSubscriptions() ) {
	    // Add query subscription tags.
	    xmlBuilder.AddAggregate( UIEVENTSUBQUERYRQ_TAG );
	    // Add subscriber id.
	    xmlBuilder.AddElement( SUBSCRIBERID_TAG, SUBSCRIBER_ID );
	    // Finish request.
	    xmlBuilder.AddEndAggregate( UIEVENTSUBQUERYRQ_TAG );
    }

    // If there are data event subscriptions,
    // add data event query request.
    if( HasDataEventSubscriptions() ) {
	    // Add query subscription tags.
	    xmlBuilder.AddAggregate( DATAEVENTSUBQUERYRQ_TAG );
	    // Add subscriber id.
	    xmlBuilder.AddElement( SUBSCRIBERID_TAG, SUBSCRIBER_ID );
	    // Finish request.
	    xmlBuilder.AddEndAggregate( DATAEVENTSUBQUERYRQ_TAG );
    }

    // Finish request.
	xmlBuilder.CreateSubscriptionXMLTrailer();

	// If any errors occurred, display error message.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

	// If no errors, return request formatted with tabs and line breaks.
	return xmlBuilder.GetXML();
}

// BuildDeleteRequest()
// Builds XML and checks for errors for DoDeleteRequest().
// Returns request XML if it could be created.
//
string CEventSubscribeDlg::BuildDeleteRequest() {
	// Begin request.
	DOMXMLBuilder xmlBuilder;
	xmlBuilder.CreateSubscriptionXMLHeader();

	// Check initial xmlBuilder access for COM and other initialization errors.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

    // If there are UI event subscriptions selected,
    // add delete request for UI events.
    if( HasUIEventSubscriptions() ) {
        // Add subscription delete tags.
        xmlBuilder.AddAggregate( SUBSCRIPTIONDELRQ_TAG );
        // Add subscriber id.
        xmlBuilder.AddElement( SUBSCRIBERID_TAG, SUBSCRIBER_ID );
        // Add subscription type.
	    xmlBuilder.AddElement( SUBSCRIPTIONTYPE_TAG, SUBSCRIPTION_TYPE_UIEVENT );
        // Close subscription delete tags.
    	xmlBuilder.AddEndAggregate( SUBSCRIPTIONDELRQ_TAG );
    }
	
    // If there are Data event subscriptions selected,
    // add delete request for Data events.
    if( HasDataEventSubscriptions() ) {
        // Add subscription delete tags.
        xmlBuilder.AddAggregate( SUBSCRIPTIONDELRQ_TAG );
        // Add subscriber id.
        xmlBuilder.AddElement( SUBSCRIBERID_TAG, SUBSCRIBER_ID );
        // Add subscription type.
	    xmlBuilder.AddElement( SUBSCRIPTIONTYPE_TAG, SUBSCRIPTION_TYPE_DATAEVENT );
        // Close subscription delete tags.
    	xmlBuilder.AddEndAggregate( SUBSCRIPTIONDELRQ_TAG );
    }
	
	// Finish request.
	xmlBuilder.CreateSubscriptionXMLTrailer();

	// If any errors occurred, display error message.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

	// If no errors, return request formatted with tabs and line breaks.
	return xmlBuilder.GetXML();
}

// BuildAddRequest()
// Builds XML and checks for errors for DoAddRequest().
// Returns request XML if it could be created.
//
string CEventSubscribeDlg::BuildAddRequest() {
	// Begin request.
	DOMXMLBuilder xmlBuilder;
	xmlBuilder.CreateSubscriptionXMLHeader();

	// Check initial xmlBuilder access for COM and other initialization errors.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

    // See if any UI events are selected.
    if( HasUIEventSubscriptions() ) {
        // If so, add UI event request.
        BuildUIEventAddRequest( &xmlBuilder );
    }

    // See if any Data events are selected.
    if( HasDataEventSubscriptions() ) {
        // If so, add Data event request.
        BuildDataEventAddRequest( &xmlBuilder );
    }

    // Add subscription closing tag.
	xmlBuilder.CreateSubscriptionXMLTrailer();

	// If any errors occurred, display error message.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

	// If no errors, return request formatted with tabs and line breaks.
	return xmlBuilder.GetXML();
}

// BuildUIEventAddRequest()
// Construct part of add request for UI events.
void CEventSubscribeDlg::BuildUIEventAddRequest( DOMXMLBuilder* pXMLBuilder )
{
	// Add UI Event opening tags.
	pXMLBuilder->AddAggregate( UIEVENTSUBADDRQ_TAG );
 	pXMLBuilder->AddAggregate( UIEVENTSUBADD_TAG );
	
	// Add subscriber id.
	pXMLBuilder->AddElement( SUBSCRIBERID_TAG, SUBSCRIBER_ID );
	
    // Add callback info.
    pXMLBuilder->AddAggregate( COMCALLBACKINFO_TAG );
    pXMLBuilder->AddElement( APPNAME_TAG, HANDLER_APP_NAME );
    pXMLBuilder->AddElement( CLSID_TAG, HANDLER_CLSID );
    pXMLBuilder->AddEndAggregate( COMCALLBACKINFO_TAG );

    // Add delivery policy.
    BOOL launchHandler = IsDlgButtonChecked( IDC_LAUNCH_HANDLER );
    if( launchHandler ) {
        pXMLBuilder->AddElement( DELIVERYPOLICY_TAG, DELIVERY_POLICY_ALWAYS );
    }
    else {
        pXMLBuilder->AddElement( DELIVERYPOLICY_TAG, DELIVERY_POLICY_ONLY_IF_RUNNING );
    }

    // Add Company File event tags.
    pXMLBuilder->AddAggregate( COMPANYFILEEVENTSUB_TAG );
    // Add company file open.
    if( IsDlgButtonChecked(IDC_FILE_OPEN) ) {
        pXMLBuilder->AddElement( COMPANYFILEEVENTOP_TAG, COMPANY_FILE_EVENT_OPEN );
    }
    // Add company file close.
    if( IsDlgButtonChecked(IDC_FILE_CLOSE) ) {
        pXMLBuilder->AddElement( COMPANYFILEEVENTOP_TAG, COMPANY_FILE_EVENT_CLOSE );
    }
    pXMLBuilder->AddEndAggregate( COMPANYFILEEVENTSUB_TAG );

    // Add UI Event closing tags.
    pXMLBuilder->AddEndAggregate( UIEVENTSUBADD_TAG );
    pXMLBuilder->AddEndAggregate( UIEVENTSUBADDRQ_TAG );
}

// BuildDataEventAddRequest()
// Construct part of add request for Data events.
void CEventSubscribeDlg::BuildDataEventAddRequest( DOMXMLBuilder* pXMLBuilder )
{
	// Add Data Event tags.
	pXMLBuilder->AddAggregate( DATAEVENTSUBADDRQ_TAG );
 	pXMLBuilder->AddAggregate( DATAEVENTSUBADD_TAG );

	// Add subscriber id.
	pXMLBuilder->AddElement( SUBSCRIBERID_TAG, SUBSCRIBER_ID );
	
	// Add callback info.
	pXMLBuilder->AddAggregate( COMCALLBACKINFO_TAG );
	pXMLBuilder->AddElement( APPNAME_TAG, HANDLER_APP_NAME );
	pXMLBuilder->AddElement( CLSID_TAG, HANDLER_CLSID );
	pXMLBuilder->AddEndAggregate( COMCALLBACKINFO_TAG );

    // Add delivery policy.
    BOOL launchHandler = IsDlgButtonChecked( IDC_LAUNCH_HANDLER );
    if( launchHandler ) {
        pXMLBuilder->AddElement( DELIVERYPOLICY_TAG, DELIVERY_POLICY_ALWAYS );
    }
    else {
        pXMLBuilder->AddElement( DELIVERYPOLICY_TAG, DELIVERY_POLICY_ONLY_IF_RUNNING );
    }

    // Set TrackLostEvents so that all lost events get tracked.
    pXMLBuilder->AddElement( TRACKLOSTEVENTS_TAG, "All" );

	// If more than one menu defined, create submenu of items.
    CListBox* pList = (CListBox*)GetDlgItem( IDC_SELECTED );
	int iCount = pList->GetCount();

	// Create XML for each defined event.
	for( int iEvent = 0; iEvent < iCount; iEvent++ ) {
        // Get event text for this event.
        CString strGetText;
        pList->GetText( iEvent, strGetText );
        string strText;
        strText = strGetText;

        // Get event info from text.
        string strSubTag, strTypeTag, strOpTag, strItem;
        BOOL bAdd, bDel, bMod, bMerge;
        bAdd = bDel = bMod = bMerge = false;
        GetEventInfo( strText,
                      strSubTag, strTypeTag, strOpTag,
                      strItem,
                      bAdd, bDel, bMod, bMerge );

        // Add tags.
        pXMLBuilder->AddAggregate( strSubTag );
        pXMLBuilder->AddElement( strTypeTag, strItem );
        if( bAdd ) {
            pXMLBuilder->AddElement( strOpTag, EVENT_OPERATION_ADD );
        }
        if( bDel ) {
            pXMLBuilder->AddElement( strOpTag, EVENT_OPERATION_DELETE );
        }
        if( bMod ) {
            pXMLBuilder->AddElement( strOpTag, EVENT_OPERATION_MODIFY );
        }
        if( bMerge ) {
            pXMLBuilder->AddElement( strOpTag, EVENT_OPERATION_MERGE );
        }
        pXMLBuilder->AddEndAggregate( strSubTag );
	}

    // Add data event closing tags.
	pXMLBuilder->AddEndAggregate( DATAEVENTSUBADD_TAG );
	pXMLBuilder->AddEndAggregate( DATAEVENTSUBADDRQ_TAG );
}


// AddEvent()
// Adds selected item to "selected" list.
void CEventSubscribeDlg::AddEvent( const string& strType,
                                   UINT  nItemList,
                                   UINT  nAdd,
                                   UINT  nMod,
                                   UINT  nDel,
                                   UINT  nMerge ) 
{
    // Get selected Transaction or List Item.
	CString strItem;
    GetDlgItemText( nItemList, strItem );
    
    // Build list of operations selected.
    CString strOperations;
    if( IsDlgButtonChecked(nAdd) ) {
        strOperations += "Add  ";
    }
    if( IsDlgButtonChecked(nMod) ) {
        strOperations += "Mod  ";
    }
    if( IsDlgButtonChecked(nDel) ) {
        strOperations += "Del  ";
    }
    if( nMerge > 0 && IsDlgButtonChecked(nMerge) ) {
        strOperations += "Merge  ";
    }

    // If no Transaction or List Item is selected,
    // display message and bail out.
    if( strItem.GetLength() == 0 ) {
        CString strMsg = "No item is selected.";
        AfxMessageBox( strMsg );
        return;
    }
    // If no operations are selected,
    // display message and bail out.
    if( strOperations.GetLength() == 0 ) {
        AfxMessageBox( "No operations are selected." );
        return;
    }

    // Build first part of text to add to "selected" list.
    CString strListItem = strType.c_str();
    strListItem += " ";
    strListItem += strItem;
    strListItem += ": ";

    // See if this item has already been added to "selected" list.
    // If it has, display message and bail out.
    CListBox* pList = (CListBox*)GetDlgItem( IDC_SELECTED );
    int count = pList->GetCount();
    for( int i = 0; i < count; i++ ) {
        CString strText;
        pList->GetText( i, strText );
        if( strText.Find(strListItem) == 0 ) {
            CString strMsg;
            strMsg += strItem;
            strMsg += " is already selected.";
            AfxMessageBox( strMsg );
            return;
        }
    }

    // Finish building text to add to "selected" list.
    strListItem += strOperations;

    // Add text to "selected" list.
    pList->AddString( strListItem );
}

// GetEventInfo()
// Extract component pieces from "selected" item string.
void CEventSubscribeDlg::GetEventInfo( const string& strText,
                                       string& strSubTag, 
                                       string& strTypeTag, 
                                       string& strOpTag,
                                       string& strItem,
                                       BOOL& bAdd, 
                                       BOOL& bDel, 
                                       BOOL& bMod, 
                                       BOOL& bMerge )
{
    // Based on type of item (List Item or Transaction), get XML
    // request tags and calculate position in string to get item 
    // description.
    int nStart = 0;
    if( strText.find(strListItemDescription) != string::npos ) {
        nStart = strListItemDescription.length() + 1;
        strSubTag  = LISTEVENTSUB_TAG;
        strTypeTag = LISTEVENTTYPE_TAG;
        strOpTag   = LISTEVENTOP_TAG;
    }
    else if( strText.find(strTransactionDescription) != string::npos ) {
        nStart = 14;
        nStart = strTransactionDescription.length() + 1;
        strSubTag  = TXNEVENTSUB_TAG;
        strTypeTag = TXNEVENTTYPE_TAG;
        strOpTag   = TXNEVENTOP_TAG;
    }

    // Extract item description from "selected" item string.
    int nEnd = strText.find( _T(":") );
    strItem = strText.substr( nStart, nEnd-nStart );

    // Determine which operations are selected for this item.
    bAdd   = strText.find( _T("Add"), nEnd ) != string::npos;
    bDel   = strText.find( _T("Del"), nEnd ) != string::npos;
    bMod   = strText.find( _T("Mod"), nEnd ) != string::npos;
    bMerge = strText.find( _T("Merge"), nEnd ) != string::npos;
}

// HasUIEventSubscriptions()
// Determine if there are any UI Event subscriptions selected.
BOOL CEventSubscribeDlg::HasUIEventSubscriptions()
{
    return IsDlgButtonChecked(IDC_FILE_OPEN) ||
           IsDlgButtonChecked(IDC_FILE_CLOSE);
}

// HasDataEventSubscriptions()
// Determine if there are any Data Event subscriptions selected.
BOOL CEventSubscribeDlg::HasDataEventSubscriptions()
{
	// If no event subscriptions defined, display error and bail out.
    CListBox* pList = (CListBox*)GetDlgItem( IDC_SELECTED );
	int iCount = pList->GetCount();
	return ( iCount > 0 );
}

// DisplayRequestError()
// Displays an error message based on a request error.
void CEventSubscribeDlg::DisplayRequestError( TCHAR* lpszMessage, 
									         qbXMLRPWrapper& requestProcessor ) {
	string strDisplay = lpszMessage;
	strDisplay += "\n\n";
	strDisplay += requestProcessor.GetErrorMsg().c_str();
	strDisplay += "\n\nThere may be additional info in the qbsdklog.txt file (in the QuickBooks executable directory).";
	AfxMessageBox( strDisplay.c_str() );
}

// DisplayXMLError()
// Displays an error message based on an XML DOM error.
void CEventSubscribeDlg::DisplayXMLError( DOMXMLBuilder& xmlBuilder ) {
	string errMsg( "Error building XML:\n\n" );
	errMsg.append( xmlBuilder.GetErrorMsg() );
	AfxMessageBox( errMsg.c_str() );
}

