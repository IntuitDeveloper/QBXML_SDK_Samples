/*---------------------------------------------------------------------------
 * FILE: MenuSubscribeDlg.cpp
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
#include "MenuSubscribe.h"
#include "MenuSubscribeDlg.h"
#include "DOMXMLBuilder.h"
#include "qbXMLTags.h"
#include "qbXMLRPWrapper.h"
#include "Utility.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

//-------------------------------------------------------------------
// CMenuSubscribeDlg dialog

// MFC-add contructor.
CMenuSubscribeDlg::CMenuSubscribeDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CMenuSubscribeDlg::IDD, pParent)
{
    // MFC-managed elements.
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	//{{AFX_DATA_INIT(CMenuSubscribeDlg)
	//}}AFX_DATA_INIT
}

// MFC-added data exchange
void CMenuSubscribeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CMenuSubscribeDlg)
	//}}AFX_DATA_MAP
}

// MFC-managed message map.
BEGIN_MESSAGE_MAP(CMenuSubscribeDlg, CDialog)
	//{{AFX_MSG_MAP(CMenuSubscribeDlg)
	ON_BN_CLICKED(IDC_ADD_SUBSCRIPTION, OnAddSubscription)
	ON_BN_CLICKED(IDC_MENU1_VISIBLE_IF, OnClickVisibleIf1)
	ON_BN_CLICKED(IDC_MENU2_VISIBLE_IF, OnClickVisibleIf2)
	ON_BN_CLICKED(IDC_MENU3_VISIBLE_IF, OnClickVisibleIf3)
	ON_BN_CLICKED(IDC_MENU1_ENABLED_IF, OnClickEnabledIf1)
	ON_BN_CLICKED(IDC_MENU2_ENABLED_IF, OnClickEnabledIf2)
	ON_BN_CLICKED(IDC_MENU3_ENABLED_IF, OnClickEnabledIf3)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


//-------------------------------------------------------------------
// CMenuSubscribeDlg message handlers

// OnInitDialog()
// Dialog initialization.
BOOL CMenuSubscribeDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	return TRUE;  // return TRUE  unless you set the focus to a control
}

// OnAddSubscription()
// Handle button click to add subscription.
void CMenuSubscribeDlg::OnAddSubscription()
{
	// If no "add to" menu selected, display error and bail out.
	string strAddToMenu = GetItemText( IDC_ADD_TO );
	if( strAddToMenu.length() == 0 ) {
		AfxMessageBox( "Please select the QuickBooks menu for adding your menu items." );
		return;
	}

	// If no menus defined, display error and bail out.
	int iCount = GetMenuCount();
	if( iCount == 0 ) {
		AfxMessageBox( "No menu items are defined." );
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
		AfxMessageBox( "Subscription added.\nRestart QuickBooks to see new menu(s)." );
	}
}

// Handlers for "VisibleIf" and "EnabledIf" buttons.
void CMenuSubscribeDlg::OnClickVisibleIf1() {
	ChangeButton( IDC_MENU1_VISIBLE_IF );
}
void CMenuSubscribeDlg::OnClickVisibleIf2() {
	ChangeButton( IDC_MENU2_VISIBLE_IF );
}
void CMenuSubscribeDlg::OnClickVisibleIf3() {
	ChangeButton( IDC_MENU3_VISIBLE_IF );
}
void CMenuSubscribeDlg::OnClickEnabledIf1() {
	ChangeButton( IDC_MENU1_ENABLED_IF );
}
void CMenuSubscribeDlg::OnClickEnabledIf2() {
	ChangeButton( IDC_MENU2_ENABLED_IF );
}
void CMenuSubscribeDlg::OnClickEnabledIf3() {
	ChangeButton( IDC_MENU3_ENABLED_IF );
}


//-------------------------------------------------------------------
// Helper methods.

// DoQueryRequest()
// Build and send request to query for a menu subscription.
//   requestProcessor = QBXML request processor used to perform requests
//   alreadyExists    = used to return whether subscription already exists
// Returns true if request succeeded; false if not.
//
BOOL CMenuSubscribeDlg::DoQueryRequest(	qbXMLRPWrapper& requestProcessor, 
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
        Utility::ShowXML( "UI Extension \"query subscription\" request",
                          "Request to check if UI subscription already exists.",
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
        Utility::ShowXML( "UI Extension \"query subscription\" response",
                          "Response from checking for existing subscription.",
                          strResponse.c_str() );
	}

    // If the response contains an error, display it and bail out.
    if( Utility::DoesResponseHaveError(QBXMLSUBMSGSRS_TAG,strResponse) ) {
        return false;
    }

	// If subscription already exists, prompt to overwrite (since there
	// can only be one UI Extension subscription per "subscriber id".)
    // We check for an existing subscription by searching for the
    // UIExtensionSubscriptionRet tag, which is returned only if a
    // matching subscription already exists.
	if( (int)strResponse.find("UIExtensionSubscriptionRet") > 0 ) {
		// Set current status;
		alreadyExists = true;
		// Display prompt.
		string strMessage = "A UI Extension subscription already exists.\n";
		strMessage += "Do you wish to overwrite it ?";
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
BOOL CMenuSubscribeDlg::DoDeleteRequest( qbXMLRPWrapper& requestProcessor ) {
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
        Utility::ShowXML( "UI Extension \"delete subscription\" request",
                          "Request to delete existing UI extension subscription.",
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
        Utility::ShowXML( "UI Extension \"delete subscription\" response",
                          "Response from deleting UI subscription.",
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
BOOL CMenuSubscribeDlg::DoAddRequest( qbXMLRPWrapper& requestProcessor ) {
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
        Utility::ShowXML( "UI Extension \"add subscription\" request",
                          "Request to add UI extension subscription.",
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
        Utility::ShowXML( "UI Extension \"add subscription\" response",
                          "Response from adding UI extension subscription.",
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
string CMenuSubscribeDlg::BuildQueryRequest() {
	// Begin request.
	DOMXMLBuilder xmlBuilder;
	xmlBuilder.CreateSubscriptionXMLHeader();

	// Check initial xmlBuilder access for COM and other initialization errors.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

	// Add UI Extension tags.
	xmlBuilder.AddAggregate( UIEXTSUBQUERYRQ_TAG );

	// Add subscriber id.
	xmlBuilder.AddElement( SUBSCRIBERID_TAG, SUBSCRIBER_ID );
	
	// Finish request.
	xmlBuilder.AddEndAggregate( UIEXTSUBQUERYRQ_TAG );
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
string CMenuSubscribeDlg::BuildDeleteRequest() {
	// Begin request.
	DOMXMLBuilder xmlBuilder;
	xmlBuilder.CreateSubscriptionXMLHeader();

	// Check initial xmlBuilder access for COM and other initialization errors.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

	// Add UI Extension tags.
	xmlBuilder.AddAggregate( SUBSCRIPTIONDELRQ_TAG );

	// Add subscriber id.
	xmlBuilder.AddElement( SUBSCRIBERID_TAG, SUBSCRIBER_ID );

	// Add subscription type.
	xmlBuilder.AddElement( SUBSCRIPTIONTYPE_TAG, SUBSCRIPTION_TYPE_UIEXTENSION );
	
	// Finish request.
	xmlBuilder.AddEndAggregate( SUBSCRIPTIONDELRQ_TAG );
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
string CMenuSubscribeDlg::BuildAddRequest() {
	// Begin request.
	DOMXMLBuilder xmlBuilder;
	xmlBuilder.CreateSubscriptionXMLHeader();

	// Check initial xmlBuilder access for COM and other initialization errors.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

	// Add UI Extension tags.
	xmlBuilder.AddAggregate( UIEXTSUBADDRQ_TAG );
	xmlBuilder.AddAggregate( UIEXTSUBADD_TAG );

	// Add subscriber id.
	xmlBuilder.AddElement( SUBSCRIBERID_TAG, SUBSCRIBER_ID );
	
	// Add callback info.
	xmlBuilder.AddAggregate( COMCALLBACKINFO_TAG );
	xmlBuilder.AddElement( APPNAME_TAG, HANDLER_APP_NAME );
	xmlBuilder.AddElement( CLSID_TAG, HANDLER_CLSID );
	xmlBuilder.AddEndAggregate( COMCALLBACKINFO_TAG );

	// Begin menu definition section.
	xmlBuilder.AddAggregate( MENUEXTSUB_TAG );
	string strAddToMenu = GetItemText( IDC_ADD_TO );
	xmlBuilder.AddElement( ADDTOMENU_TAG, strAddToMenu );

	// If more than one menu defined, create submenu of items.
	int iCount = GetMenuCount();
	if( iCount > 1 ) {
		xmlBuilder.AddAggregate( SUBMENU_TAG );
	}

	// Create XML for each defined menu item.
	for( int iMenu = 0; iMenu < 3; iMenu++ ) {
		// Menu text.
		string strMenuText = GetMenuText( iMenu );

		// If menu text defined, add menu item.
		if( strMenuText.length() > 0 ) {
			// Begin this menu item definition.
			xmlBuilder.AddAggregate( MENUITEM_TAG );
			xmlBuilder.AddElement( MENUTEXT_TAG, strMenuText );

			// Menu tag (menu_1,menu_2,etc.) This gets echoed back in
			// event XML when menu item is clicked by user.
			string strEventTag = "menu_";
			char lpszID[2];
			lpszID[0] = (char)('1'+iMenu);
			lpszID[1] = 0;
			strEventTag.append( lpszID );
			xmlBuilder.AddElement( EVENTTAG_TAG, strEventTag );

			// Display conditions.
			string strVisibleCondition = GetVisibleCondition( iMenu );
			string strEnabledCondition = GetEnabledCondition( iMenu );
			if( strVisibleCondition.length() > 0 || strEnabledCondition.length() > 0 ) {
				xmlBuilder.AddAggregate( DISPLAYCONDITION_TAG );
				// Visible condition.
				if( strVisibleCondition.length() > 0 ) {
					BOOL bVisibleIf = GetVisibleIf( iMenu );
					if( bVisibleIf ) {
						xmlBuilder.AddElement( VISIBLEIF_TAG, strVisibleCondition );
					}
					else {
						xmlBuilder.AddElement( VISIBLEIFNOT_TAG, strVisibleCondition );
					}
				}
				// Enabled condition.
				if( strEnabledCondition.length() > 0 ) {
					BOOL bEnabledIf = GetEnabledIf( iMenu );
					if( bEnabledIf ) {
						xmlBuilder.AddElement( ENABLEDIF_TAG, strEnabledCondition );
					}
					else {
						xmlBuilder.AddElement( ENABLEDIFNOT_TAG, strEnabledCondition );
					}
				}
				xmlBuilder.AddEndAggregate( DISPLAYCONDITION_TAG );
			}

			// End menu item aggregate.
			xmlBuilder.AddEndAggregate( MENUITEM_TAG );
		}
	}

	// If more than one menu defined, end submenu aggregate.
	if( iCount > 1 ) {
		xmlBuilder.AddEndAggregate( SUBMENU_TAG );
	}

	// Finish request.
	xmlBuilder.AddEndAggregate( MENUEXTSUB_TAG );

	xmlBuilder.AddEndAggregate( UIEXTSUBADD_TAG );
	xmlBuilder.AddEndAggregate( UIEXTSUBADDRQ_TAG );
	xmlBuilder.CreateSubscriptionXMLTrailer();

	// If any errors occurred, display error message.
	if( xmlBuilder.GetHasError() ) {
		DisplayXMLError( xmlBuilder );
		return "";
	}

	// If no errors, return request formatted with tabs and line breaks.
	return xmlBuilder.GetXML();
}

// GetItemText()
// Return string object containing text value of given dialog control.
string CMenuSubscribeDlg::GetItemText( int id ) {
	string strMenu;

	TCHAR lpszMenu[100];
	if( GetDlgItemText( id, lpszMenu, 99 ) > 0 ) {
		strMenu.assign( lpszMenu );
	}
	else {
		strMenu.assign( "" );
	}

	return strMenu;
}

// GetMenuCount()
// Return number of menu items that are defined.
int CMenuSubscribeDlg::GetMenuCount() {
	int iCount = 0;

    // A menu item is defined if menu text is entered.
	for( int i = 0; i < 3; i++ ) {
		if( GetMenuText(i).length() > 0 ) iCount++;
	}

	return iCount;
}

// GetMenuText()
// Returns menu text for given menu number.
string CMenuSubscribeDlg::GetMenuText( int iMenu ) {
	int id = IDC_MENU_BASE + ( iMenu * 10 );
	return GetItemText( id );
}

// GetVisibleIf()
// Returns VisibleIf value for given menu number.
//  Indicates whether selected condition should be "reversed".
BOOL CMenuSubscribeDlg::GetVisibleIf( int iMenu ) {
	int id = IDC_MENU_BASE + ( iMenu * 10 ) + 1;
	return !IsDlgButtonChecked( id );
}

// GetVisibleCondition()
// Returns visible condition for given menu number.
string CMenuSubscribeDlg::GetVisibleCondition( int iMenu ) {
	int id = IDC_MENU_BASE + ( iMenu * 10 ) + 2;
	return GetItemText( id );
}

// GetEnabledIf()
// Returns EnabledIf value for given menu number.
//  Indicates whether selected condition should be "reversed".
BOOL CMenuSubscribeDlg::GetEnabledIf( int iMenu ) {
	int id = IDC_MENU_BASE + ( iMenu * 10 ) + 3;
	return !IsDlgButtonChecked( id );
}

// GetEnabledCondition()
// Returns enabled condition for given menu number.
string CMenuSubscribeDlg::GetEnabledCondition( int iMenu ) {
	int id = IDC_MENU_BASE + ( iMenu * 10 ) + 4;
	return GetItemText( id );
}

// ChangeButton()
// Handles toggling between "If" and "If Not" when button is clicked.
void CMenuSubscribeDlg::ChangeButton( int id ) {
	BOOL isChecked = IsDlgButtonChecked( id );
	if( isChecked ) {
		SetDlgItemText( id, "If Not" );
	}
	else {
		SetDlgItemText( id, "If" );
	}
}

// DisplayRequestError()
// Displays an error message based on a request error.
void CMenuSubscribeDlg::DisplayRequestError( TCHAR* lpszMessage, 
									         qbXMLRPWrapper& requestProcessor ) {
	string strDisplay = lpszMessage;
	strDisplay += "\n\n";
	strDisplay += requestProcessor.GetErrorMsg().c_str();
	strDisplay += "\n\nThere may be additional info in the qbsdklog.txt file (in the QuickBooks executable directory).";
	AfxMessageBox( strDisplay.c_str() );
}

// DisplayXMLError()
// Displays an error message based on an XML DOM error.
void CMenuSubscribeDlg::DisplayXMLError( DOMXMLBuilder& xmlBuilder ) {
	string errMsg( "Error building XML:\n\n" );
	errMsg.append( xmlBuilder.GetErrorMsg() );
	AfxMessageBox( errMsg.c_str() );
}

