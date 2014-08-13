/*---------------------------------------------------------------------------
 * FILE: Utility.cpp
 *
 * Description:
 * Implementation BSTR convertion methods and get Node attributes and
 * value methods.
 *
 * Created On: 11/11/2001
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#include "stdafx.h"
#include <time.h>
#include "Utility.h"
#include "qbXMLTags.h"
#include "ShowXMLDlg.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

Utility::Utility()
{

}

Utility::~Utility()
{

}

// initialize static data
static unsigned int s_randSeed = 0;

/*---------------------------------------------------------------------------
 * BSTRToString
 * Static method, it converts BSTR to string.
 *
 */
void Utility::BSTRToString(BSTR inVal, string &outVal)
{
  TCHAR *tmpVal = _com_util::ConvertBSTRToString(inVal);

  if (tmpVal == NULL) {
    outVal = "";
  }
  else {
    outVal = tmpVal;
    // com_util allocates the buffer, we have to free it.
    delete [] tmpVal;
  }
}

/*---------------------------------------------------------------------------
 * loadXML
 * Static method, it loads XML string to XML DOM obj.
 *
 */
MSXML2::IXMLDOMDocument *Utility::LoadXML(const string &xmlStr, string &reason)
{
  HRESULT hr;

  MSXML2::IXMLDOMDocument *xmlDocPtr = NULL;
  hr = CoCreateInstance(MSXML2::CLSID_DOMDocument,
                        NULL,
                        CLSCTX_INPROC_SERVER,
                        MSXML2::IID_IXMLDOMDocument,
                        reinterpret_cast<void**>(&xmlDocPtr));

  if (SUCCEEDED(hr)) {

    CComBSTR cbstrXML(xmlStr.c_str()); // for BSTR
    VARIANT_BOOL vBool;

    // load the xml doc
    hr = xmlDocPtr ->loadXML(cbstrXML.m_str, &vBool);

    if (FAILED(hr)) {

      // load xml doc failed
      // get the parse error reason
      MSXML2::IXMLDOMParseError *pError = NULL;

      xmlDocPtr ->get_parseError(&pError);

      CComBSTR bstrReason;

      pError ->get_reason(&bstrReason);

      BSTRToString(bstrReason, reason);

      // Release the Document object
      xmlDocPtr ->Release();

      return NULL;
    }
  }
  else { // CreateInstance failed
    reason = "Instantiate MSXML DOM OBJ failed";
    return NULL;
  }

  return xmlDocPtr;
}


/*---------------------------------------------------------------------------
 * GetAttrValue
 * Static method, to get the attribute value from a IXMLDOMNode
 * by passing the attribute name.
 *
 */

HRESULT Utility::GetAttrValue(MSXML2::IXMLDOMNamedNodeMap *nodeMap,
                              const string &attrName,
                              string &attrValue)
{
  CComBSTR bName(attrName.c_str());

  MSXML2::IXMLDOMNode *attrNode = NULL;
  HRESULT hr = nodeMap ->getNamedItem(bName.m_str, &attrNode);
  VARIANT vValue;

  if (SUCCEEDED (hr)) {
    attrNode ->get_nodeValue(&vValue);

    // convert to TCHAR
    TCHAR *tmpVal = _com_util::ConvertBSTRToString(vValue.bstrVal);
    attrValue = tmpVal;

    delete [] tmpVal;
    attrNode ->Release();
  }

  return hr;
}


/*---------------------------------------------------------------------------
 * GetNodeValue
 * static method, to get the Node value from a IXMLDOMNamedNodeMap
 * by passing the attribute name.
 *
 */
HRESULT Utility::GetNodeValue(MSXML2::IXMLDOMNode *node,
                              string &nodeValue)
{
  HRESULT hr = S_OK;

  if( node == NULL) {
    return hr;
  }

  CComBSTR bstrValue;
  hr = node ->get_text(&bstrValue);

  // convert to TCHAR
  if(SUCCEEDED(hr)){
    TCHAR *nodeValuePtr = NULL;
    nodeValuePtr = _com_util::ConvertBSTRToString(bstrValue);

    if (nodeValuePtr != NULL){
      nodeValue = nodeValuePtr;
      delete [] nodeValuePtr;
    }
  }

  node ->Release();

  return hr;
}

HRESULT Utility::GetNodeValue(MSXML2::IXMLDOMNode *node,
                              TCHAR *nodeValue)
{
  HRESULT hr;
  string tmpVal;

  hr = GetNodeValue(node, tmpVal);

  if(SUCCEEDED(hr) && !tmpVal.empty()) {
    strcpy(nodeValue, tmpVal.c_str());
  }

  return hr;
}

/*---------------------------------------------------------------------------
 * DoesResponseHaveError
 * static method, to check for error in response and display message if 
 * there is an error.
 * Returns true if error found.
 *
 */
BOOL Utility::DoesResponseHaveError( LPCTSTR lpszTag, 
                                     const string& strResponse ) {
    // Parse response into a DOM object.
    string strReason;
    MSXML2::IXMLDOMDocument *pXMLDocPtr = Utility::LoadXML( strResponse, strReason );
    if( pXMLDocPtr == NULL ) {
        string strMessage = "Error parsing response XML:\n\n";
        strMessage += strReason;
        AfxMessageBox( strMessage.c_str() );
        return true;
    }

    // Look for response object.
    MSXML2::IXMLDOMNodeList* pResultList = NULL;
    CComBSTR tagName( lpszTag );
    HRESULT hr = pXMLDocPtr->getElementsByTagName( tagName, &pResultList );
    if( FAILED(hr) || pResultList == NULL ) {
        AfxMessageBox( "Error parsing response XML:\n\nUnable to find response aggregate" );
        return true;
    }
    
    // Make sure we got a response object.
    long tagCount;
    hr = pResultList->get_length( &tagCount );
    if( FAILED(hr) || tagCount < 1 ) {
        AfxMessageBox( "Error parsing response XML:\n\nResponse aggregate not found" );
        return true;
    }

    // Get the response object.
    MSXML2::IXMLDOMNode *pListItem = NULL;
    hr = pResultList->get_item( 0, &pListItem );
    if( FAILED(hr) || pListItem == NULL ) {
        AfxMessageBox( "Error parsing response XML:\n\nUnable to read response aggregate" );
        return true;
    }

    // Get the actual response node from the response object.
    MSXML2::IXMLDOMNode *pResponse = NULL;
    hr = pListItem->get_firstChild( &pResponse );
    if( FAILED(hr) || pListItem == NULL ) {
        AfxMessageBox( "Error parsing response XML:\n\nUnable to read response" );
        return true;
    }

    // Get response attributes.
    MSXML2::IXMLDOMNamedNodeMap *pAttributes = NULL;
    hr = pResponse->get_attributes( &pAttributes );
    if( FAILED(hr) || pAttributes == NULL ) {
        AfxMessageBox( "Error parsing response XML:\n\nUnable to get response attributes" );
        return true;
    }

    // Get response status info.
    string statusCode;
    string statusSeverity;
    string statusMessage;
    hr = Utility::GetAttrValue( pAttributes, ATTR_STATUSCODE_TAG, statusCode );
    if( SUCCEEDED(hr) ) {
        hr = Utility::GetAttrValue( pAttributes, ATTR_STATUSSEVERITY_TAG, statusSeverity );
    }
    if( SUCCEEDED(hr) ) {
        hr = Utility::GetAttrValue( pAttributes, ATTR_STATUSMESSAGE_TAG, statusMessage );
    }
    if( FAILED(hr) ) {
        AfxMessageBox( "Error parsing response XML:\n\nUnable to read response status" );
        return true;
    }
    
    // If the response contains an error, display error message.
    if( statusSeverity == "Error" ) {
        string strMessage = "Error with request:\n\n";
        strMessage += statusMessage;
        AfxMessageBox( strMessage.c_str() );
        return true;
    }

    return false;
}


/*---------------------------------------------------------------------------
 * ShowXML
 * static method, to show XML requests and responses in a dialog, which 
 * allows copying of the XML.
 *
 */
void Utility::ShowXML( LPCTSTR lpszTitle, LPCTSTR lpszHeader, LPCTSTR lpszXML ) {
    CShowXMLDlg dlg;
    dlg.SetTitle( lpszTitle );
    dlg.SetHeader( lpszHeader );
    dlg.SetXML( lpszXML );
    dlg.DoModal();
}

void Utility::GetRequestID(std::string &rID)
{
  TCHAR *tmpID = Utility::GetRequestID();

  rID = tmpID;
  delete [] tmpID;
}

// using rand() to generate a request ID
TCHAR * Utility::GetRequestID()
{
  // Seed the random number generator on first access.
  if (s_randSeed == 0) {
    s_randSeed = (unsigned int) time( NULL );
    srand(s_randSeed);
  }

  int randNum = rand();

  const int requestID_LEN = 50;
  const TCHAR * requestIDNM = "requestID";

  TCHAR *thisRequestID = new TCHAR[requestID_LEN + strlen(requestIDNM) + 1];

  sprintf(thisRequestID, "%s_%d", requestIDNM, randNum);

  return thisRequestID;
}