/*---------------------------------------------------------------------------
 * FILE: DOMXMLBuilder.cpp
 *
 * Description:
 * DOMXMLBuilder Class Implementation. It provides the basic methods to
 * create XML doc by using MSXML. The methods include:
 *                1. CreateXMLHeader:   Add top of the qbXML request
 *                2. CreateXMLTrailer:  Add bottom of the qbXML request
 *                3. AddAggregate:      Add a new aggregate
 *                4. AddEndAggregate:   Add a closing tag for an aggregate
 *                5. AddElement:        Add a new element
 *
 * There are also accessors to get the resulting XML and error status
 * and message:
 *                6. GetXML:            Get the XML we built
 *                7. GetHasError:       Returns true if there was an error
 *                8. GetErrorMsg:       Returns the corresponding message for an error
 *
 * Created On: 11/11/2001
 * Updated to SDK 2.0: 08/08/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */
#include "stdafx.h"
#include "DOMXMLBuilder.h"

#include "qbXMLTags.h"
#include "Utility.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DOMXMLBuilder::DOMXMLBuilder()
  :m_IsTopLevel(true),
   m_HasError(false),
   m_XMLDoc(NULL),
   m_CurrentNode(NULL)
{

  m_XML = "";
  m_Prolog = "";
  m_ErrorMsg = "OK";

}

DOMXMLBuilder::~DOMXMLBuilder()
{
  // Release COM objects
  if (m_XMLDoc != NULL) {
    m_XMLDoc ->Release();
    m_XMLDoc = NULL;
  }
  if (m_CurrentNode != NULL) {
    m_CurrentNode ->Release();
    m_CurrentNode = NULL;
  }
}


/*---------------------------------------------------------------------------
 * InstantiateXMLDomDocument
 * Creates an instance of the MSXML DOM Document so we can start
 * building XML.
 *
 */
bool DOMXMLBuilder::InstantiateXMLDomDocument()
{
  HRESULT hr;
  hr = CoCreateInstance(MSXML2::CLSID_DOMDocument40,
                        NULL,
                        CLSCTX_INPROC_SERVER,
                        MSXML2::IID_IXMLDOMDocument,
                        reinterpret_cast<void**>(&m_XMLDoc));


  // COM Create Instance failed
  if (FAILED(hr)){
    TCHAR tmpMsg[1024];
    sprintf (tmpMsg, "Failed to instantiate the MSXML2 DOM Document.  Error: 0x%lx.  Perhaps msxml4.dll is missing?", hr);
    m_ErrorMsg = tmpMsg;
    m_HasError = true;
    return false;
  }

  return true;
}


/*---------------------------------------------------------------------------
 * GetXML
 *
 * You should only call this method when you are done building the XML.
 *
 */
string &DOMXMLBuilder::GetXML()
{
  if (m_XML != "") {
    return m_XML;
  }

  if (!m_HasError) {
    CComBSTR bstrXMLStr;

    // Get XML string from MSXML DOM document
    m_XMLDoc ->get_xml(&bstrXMLStr);

    // concatenate prolog and the XML string
    m_XML = m_Prolog;

    string xmlBody;
    Utility::BSTRToString(bstrXMLStr, xmlBody);
    m_XML += xmlBody;
  }

  return m_XML;
}


/*---------------------------------------------------------------------------
 * AddAggregate
 * Add XML Aggregate start tag.  You can specify attributes here, if any.
 *
 */
void DOMXMLBuilder::AddAggregate(const string &inputTag,
                                 const string attNameArray[],
                                 const string attValueArray[],
                                 int   attCount)
{
  HRESULT hr;
  CComBSTR tag(inputTag.c_str());

  // new element
  MSXML2::IXMLDOMElement *newElement = NULL;
  MSXML2::IXMLDOMNode    *tmpNode = NULL;

  // Instantiate MSXML if we haven't already
  if (m_XMLDoc == NULL) {
    if (!InstantiateXMLDomDocument()) {
      return;
    }
  }

  // Create a new element
  hr = m_XMLDoc ->createElement(tag.m_str, &newElement);

  if (FAILED(hr)) { // failed
    m_HasError = true;
    m_ErrorMsg = "createElement for new aggregate failed";
    return;
  }

  if (m_IsTopLevel) {
    m_IsTopLevel = false;

    // append this new node to the document root
    hr = m_XMLDoc ->appendChild(newElement, &tmpNode);
  }
  else {
    // append this new node to the current node
    hr = m_CurrentNode ->appendChild(newElement, &tmpNode);
  }

  if (FAILED(hr)) { // failed
    m_HasError = true;
    m_ErrorMsg = "appendChild for new aggregate failed";
    newElement ->Release();
    return;
  }

  // We don't need the tmpNode that appendChild returns.
  tmpNode ->Release();

  // set current node
  if (m_CurrentNode != NULL) {
    m_CurrentNode ->Release();
  }
  m_CurrentNode = newElement;

  // set attribute list here
  for (int i = 0; i < attCount; ++i) {
    CComBSTR name(attNameArray[i].c_str());
    _variant_t value(attValueArray[i].c_str()); // VARIANT
    hr = newElement ->setAttribute(name.m_str, value);

    if (FAILED(hr)) {
      m_HasError = true;
      m_ErrorMsg = "setAttribute on new aggregate failed";
      return;
    }
  }
}


/*---------------------------------------------------------------------------
 * AddEndAggregate
 * Add XML Aggregate end tag, and move the current node to parent node
 *
 */
void DOMXMLBuilder::AddEndAggregate(const string &inputTag)
{
  if (m_CurrentNode == NULL) {
    m_HasError = true;
    m_ErrorMsg = "Cannot add EndAggregate.  Current node is NULL!";
    return;
  }

  // move up the current node
  MSXML2::IXMLDOMNode *pNode;
  m_CurrentNode ->get_parentNode(&pNode);
  m_CurrentNode ->Release();
  m_CurrentNode = pNode;

}


/*---------------------------------------------------------------------------
 * AddElement
 * Add XML Element
 *
 */
void DOMXMLBuilder::AddElement(const string &inTag,
                               const string &value) {

  HRESULT hr;

  MSXML2::IXMLDOMElement *newElement = NULL;

  if (m_CurrentNode == NULL) {
    m_HasError = true;
    m_ErrorMsg = "Cannot add element.  Current node is NULL!";
    return;
  }

  // Create new element
  CComBSTR tag(inTag.c_str());
  hr = m_XMLDoc ->createElement(tag.m_str, &newElement);

  if (FAILED(hr)){ // create element failed
    m_HasError = true;
    m_ErrorMsg = "createElement for new element failed";
    return;
  }

  CComBSTR bstrValue(value.c_str());

  // Set the value
  newElement ->put_text(bstrValue);

  // Append it to the current node
  MSXML2::IXMLDOMNode *tmpNode;

  hr = m_CurrentNode ->appendChild(newElement, &tmpNode);
  if (FAILED(hr)) { // failed
    m_HasError = true;
    m_ErrorMsg = "appendChild on new element failed";
    newElement ->Release();
    return;
  }

  tmpNode ->Release();
  newElement ->Release();
}


/*---------------------------------------------------------------------------
 * CreateXMLHeader
 * Create Top Level prolog and aggregates of the request XML
 *
 */
void DOMXMLBuilder::CreateXMLHeader()
{
  // Instantiate MSXML if we haven't already
  if (m_XMLDoc == NULL) {
    if (!InstantiateXMLDomDocument()) {
      return;
    }
  }

  // Top Level
  AddTagTopLevel ("?xml version=\"1.0\" ?");

  // Use new qbxml version tag instead of DOCTYPE, although
  // DOCTYPE can also be used
  AddTagTopLevel ("?qbxml version=\"2.0\" ?");

  // <QBXML>
  AddAggregate (QBXML_TAG);

  // <QBXMLMsgsRq onError=stopOnError>
  string attrsName[1];
  string attrsValue[1];

  int attrIndex = 0;

  attrsName[attrIndex] = ATTR_ONERROR_TAG;
  attrsValue[attrIndex] = "stopOnError";
  attrIndex++;

  AddAggregate (QBXMLMSGSRQ_TAG, attrsName, attrsValue, attrIndex);
}


/*---------------------------------------------------------------------------
 * CreateXMLTrailer
 * Create the trailing end aggregates of the request XML
 *
 */
void DOMXMLBuilder::CreateXMLTrailer()
{
  AddEndAggregate (QBXMLMSGSRQ_TAG);
  AddEndAggregate (QBXML_TAG);
}


/*---------------------------------------------------------------------------
 * AddTagTopLevel
 * Add XML top level tags
 *
 */
void DOMXMLBuilder::AddTagTopLevel(const TCHAR *inputStr)
{
  if (inputStr != NULL) {
    m_Prolog += "<";
    m_Prolog += inputStr;
    m_Prolog += ">\n";
  }
}

