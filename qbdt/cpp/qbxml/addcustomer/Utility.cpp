/*---------------------------------------------------------------------------
 * FILE: Utility.cpp
 *
 * Description:
 * Implementation BSTR convertion methods and get Node attributes and
 * value methods.
 *
 * Created On: 11/07/2001
 * Updated to QBXML 2.0 August 2002
 *
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#include "stdafx.h"
#include <time.h>
#include "Utility.h"

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
  hr = CoCreateInstance(MSXML2::CLSID_DOMDocument40,
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
    s_randSeed = time( NULL );
    srand(s_randSeed);
  }

  int randNum = rand();

  const int requestID_LEN = 50;
  const TCHAR * requestIDNM = "requestID";

  TCHAR *thisRequestID = new TCHAR[requestID_LEN + strlen(requestIDNM) + 1];

  sprintf(thisRequestID, "%s_%d", requestIDNM, randNum);

  return thisRequestID;
}