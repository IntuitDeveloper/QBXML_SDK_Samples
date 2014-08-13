/*---------------------------------------------------------------------------
 * FILE: QBSDKCallback.cpp
 *
 * Description:
 * Implements callback for QB event notification.
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
#include "EventHandler.h"
#include "QBSDKCallback.h"
#include "MainFrm.h"
#include "Utility.h"
#include "qbXMLTags.h"
#include "DOMXMLBuilder.h"
#include "qbXMLRPWrapper.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
//

// This is the method called by QBSDK when an event is fired
// in QuickBooks and you have subscribed to that event.
STDMETHODIMP CQBSDKCallback::inform(BSTR eventXML) {
    HRESULT hr = S_OK;

    // NOTE: This test QBSDK callback method uses message boxes.
    // A real application should not do this or any other "blocking"
    // actions in the callback, because the callback should return
    // as soon as possible. This is especially important when handling
    // UIExtension events, because additional clicks on your menu will
    // be ignored until your callback returns, making your menu item
    // non-responsive. If you need to perform blocking operations, 
    // rather than performing them directly in your callback, you
    // should post a user-defined message that you handle later and
    // perform the blocking operations there. This will allow the 
    // callback to return quickly.

    // Get app settings.
    CEventHandlerApp* pApp = (CEventHandlerApp*)AfxGetApp();
    BOOL launchApp = pApp->GetProfileInt(   pApp->lpszProfileSection,
                                            pApp->lpszProfileLaunch,
                                            1 );
    BOOL recoveryAlert = pApp->GetProfileInt(   pApp->lpszProfileSection,
                                                pApp->lpszProfileRecoveryAlert,
                                                1 );
    CString strOutputPath = pApp->GetProfileString( pApp->lpszProfileSection,
                                                    pApp->lpszProfileOutputPath,
                                                    _T("") );
    // If app's window already exists, show it.
    CMainFrame* pFrame = (CMainFrame*)pApp->m_pMainWnd;
    if( pFrame != NULL ) {
        pFrame->ShowWindow( SW_SHOWNORMAL );
        pFrame->BringToForeground();
    }
    // If app's window does not yet exist, create and show it
    // if that is the current setting. Otherwise, the XML will
    // be displayed in a message box.
    else if( launchApp ) {
        pApp->CreateMainWindow();
        pFrame = (CMainFrame*)AfxGetMainWnd();
    }


    // Format the received event XML with tabs and line breaks.
    string strEventXML;
    Utility::BSTRToString( eventXML, strEventXML );
    string strFormattedXML = DOMXMLBuilder::FormatXML( strEventXML );

    // If an output path is specified, write XML to file.
    if( strOutputPath.GetLength() > 0 ) {
        WriteToFile( strOutputPath, 
                     strFormattedXML.c_str() );
    }

    // If the app's window exists, show XML in main app window.
    if( pFrame ) {
        pFrame->ShowEventXML( strFormattedXML );
    }
    // If we couldn't get or create app's window, display
    // XML in a message box.
    else {
        AfxMessageBox( strFormattedXML.c_str() );
    }

    // If option is set to show recovery alert, see if there is
    // a recovery time stamp in event XML.
    if( recoveryAlert ) {
        // If there is a recovery time stamp, we've missed an event.
        if( (int)strEventXML.find(DATAEVENTRECOVERYTIME_TAG) > 0 ) {
            string strCompanyFile = GetTagValue(strEventXML,COMPANYFILEPATH_TAG);
            RecoveryAlert( strCompanyFile );
        }
    }

    return hr;
}

// Write event XML to output file.
void CQBSDKCallback::WriteToFile( LPCTSTR lpszPath, LPCTSTR lpszXML   )
{
    // Get output file name.
    CString strOutputPath = lpszPath;
    // If path doesn't end with backslash, add it.
    if( '\\' != strOutputPath.GetAt(strOutputPath.GetLength()-1) ) {
        strOutputPath += "\\";
    }
    // Add file name.
    strOutputPath += "EventXML.txt";

    // Append XML to file.
    CStdioFile file;
    int nOpenFlags =    CFile::modeCreate | 
                        CFile::modeNoTruncate | 
                        CFile::modeReadWrite | 
                        CFile::typeText;
	if( file.Open( strOutputPath, nOpenFlags ) ) {
		ULONGLONG len = file.SeekToEnd();
		if( len > 0 ) {
			file.WriteString( "\n--------------------------------------------------------\n" );
		}
		file.WriteString( lpszXML );
		file.Flush();
		file.Close();
	}
}

// Show message that events have been lost and prompt if
// recovery time-stamp should be cleared.
void CQBSDKCallback::RecoveryAlert( const string& strCompanyFile ) 
{
    // Prompt if time stamp should be cleared.
    CString strMessage = "The event XML contains a data recovery time-stamp.\n";
    strMessage += "This value is provided when events have not been handled, ";
    strMessage += "usually because they occurred on another machine.\n\n";
    strMessage += "Clear data recovery time-stamp ?";

    // If not clearing time-stamp, bail out.
    if( IDYES != AfxMessageBox(strMessage,MB_YESNO|MB_ICONQUESTION) ) {
        return;
    }

    // Send request to clear timestamp.
    // In a real application, we would first make requests to get the
    // data that we missed and were interested in. For example, if our
    // app synchronized invoices with QB, we would query for invoices
    // that had changed or been deleted, and update our data with that 
    // info. The invoice query would use the recovery timestamp value 
    // from the event notification that we're currently handling, and
    // use that value for the FromModifiedDate and FromDeletedDate 
    // values in the invoice query.

	// Begin request.
   	DOMXMLBuilder xmlBuilder;
	xmlBuilder.CreateXMLHeader();

    // Add time-stamp clearing tags.
	xmlBuilder.AddAggregate( DATAEVENTRECOVERYINFODELRQ_TAG );
	xmlBuilder.AddElement( SUBSCRIBERID_TAG, SUBSCRIBER_ID );
	xmlBuilder.AddEndAggregate( DATAEVENTRECOVERYINFODELRQ_TAG );

	// Finish request.
	xmlBuilder.CreateXMLTrailer();

	// Check initial xmlBuilder access for COM and other initialization errors.
	if( xmlBuilder.GetHasError() ) {
        AfxMessageBox( "Error building request.\nTime-stamp not cleared." );
        return;
	}

    // Get request as a string.
    string strRequest = xmlBuilder.GetXML();

    // Display request.
    string strFormattedRequest = DOMXMLBuilder::FormatXML( strRequest );
    strMessage = "Sending request to clear recovery time-stamp:\n\n";
    strMessage += strFormattedRequest.c_str();
    // If user cancels, bail out.
    if( IDCANCEL == AfxMessageBox( strMessage, MB_OKCANCEL ) ) {
        return;
    }

    CWaitCursor waitCursor;
   	qbXMLRPWrapper requestProcessor;

    // Open SDK connection.
    if( !requestProcessor.OpenCompanyFile( strCompanyFile ) ) {
        strMessage = "Error connecting to company file:\n  ";
        strMessage += strCompanyFile.c_str();
        strMessage += "\n\nTime-stamp not cleared.";
        AfxMessageBox( strMessage );
        return;
    }

    // Send the request.
	string strResponse = requestProcessor.ProcessRequest( strRequest );

    // If an error occurred, display message.
	if( requestProcessor.GetHasError() ) {
        strMessage = "An error occurred with the request to clear the recovery time-stamp.\n";
        strMessage += "Check \"qbsdklog.txt\" file (in QuickBooks executable directory) for additional info.\n\n";
        strMessage += "Time-stamp was not cleared.";
        AfxMessageBox( strMessage );
    }
    // If request succeeded, display response.
    else {
        string strFormattedResponse = DOMXMLBuilder::FormatXML( strResponse );
        strMessage = "Recovery time-stamp cleared:\n\n";
        strMessage += strFormattedResponse.c_str();
        AfxMessageBox( strMessage );
    }

}

// "Quick and dirty" method to get a tag value out of event XML.
string CQBSDKCallback::GetTagValue( const string& strXML, const string& strTag )
{
    string strTagValue = "";

    // Build opening tag for value we're looking for.
    string strOpenTag = "<";
    strOpenTag += strTag;
    strOpenTag += ">";

    // Build closing tag for value we're looking for.
    string strCloseTag = "</";
    strCloseTag += strTag;
    strCloseTag += ">";

    // Find opening and closing tags.
    int posOpen = 0;
    int posClose = 0;
    posOpen = strXML.find( strOpenTag );
    if( posOpen > 0 ) {
        posOpen += strOpenTag.length();
        posClose = strXML.find( strCloseTag, posOpen );
    }

    // Extract value.
    if( posClose > 0 ) {
        strTagValue = strXML.substr( posOpen, posClose-posOpen );
    }

    return strTagValue;
}
