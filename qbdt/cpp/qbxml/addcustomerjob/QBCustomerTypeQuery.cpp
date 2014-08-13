/*---------------------------------------------------------------------------
 * FILE: QBCustomerTypeQuery.cpp
 *
 * Description:
 * Implementation of CustomerType Query operation.
 *
 * Created On: 11/11/2001
 * Updated to SDK 2.0: 08/08/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "QBCustomerTypeQuery.h"
#include "qbXMLRPWrapper.h"
#include "Utility.h"
#include "qbXMLTags.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

QBCustomerTypeQuery::QBCustomerTypeQuery(const string &companyFileName,
                                         qbXMLRPWrapper *pQbxmlrpWrapper)
{
  m_CompanyFileName = companyFileName;
  m_pQbxmlrpWrapper = pQbxmlrpWrapper;
}

QBCustomerTypeQuery::~QBCustomerTypeQuery()
{
  // clean up the customer Type list
  QBCustTypeRetVec::iterator it;

  for (it = m_CustTypeRetVec.begin(); it != m_CustTypeRetVec.end(); it++) {
    QBCustomerTypeRet *custTypeRet = *it;
    delete custTypeRet;
  }
}


/*---------------------------------------------------------------------------
 * GetCustomerTypeListFromQB
 * Get all customer types from QuickBooks.  It creates the request XML,
 * sends to QuickBooks, and parses the response XML. It uses the base class
 * method PerformOperation().
 *
 */
bool QBCustomerTypeQuery::GetCustomerTypeListFromQB()
{
  return PerformOperation();
}


/*---------------------------------------------------------------------------
 * CreateXMLBody
 * Create CustomerTypeQueryRq XML fragment.
 *
 */
bool QBCustomerTypeQuery::CreateXMLBody()
{
  string attrsName[1];
  string attrsValue[1];

  int attrIndex = 0;

  string currRequestID;
  Utility::GetRequestID (currRequestID);
  attrsName[attrIndex] = ATTR_REQUESTID_TAG;
  attrsValue[attrIndex] = currRequestID;

  attrIndex++;

  // the query to get all customer types
  m_Builder -> AddAggregate (CUSTOMERTYPEQUERYRQ_TAG, attrsName, attrsValue, attrIndex);
  m_Builder -> AddEndAggregate (CUSTOMERTYPEQUERYRQ_TAG);

  if (m_Builder ->GetHasError()) {
    m_ErrorMsg = m_Builder ->GetErrorMsg();
    return false;
  }

  return true;
}


/*---------------------------------------------------------------------------
 * ParseXML
 * Uses the MSXML 4.0 DOM obj to parse the response XML.
 *
 */
bool QBCustomerTypeQuery::ParseXML()
{
  // Create xmlDoc Obj

  MSXML2::IXMLDOMDocument *xmlDoc = NULL;
  string reason;

  xmlDoc = Utility::LoadXML (m_ResponseXML, reason);

  if (xmlDoc == NULL) {
    m_ErrorMsg = "Load XML Doc failed: " + reason;
    return false;
  }

  // walk through the DOM tree
  // to find the CustomerTypeQueryRs Element and
  // its statusCode, statusSeverity, statusMessage attributes

  MSXML2::IXMLDOMNodeList *nodeList = NULL;

  CComBSTR tagName (CUSTOMERTYPEQUERYRS_TAG);

  // Get the list of elements with the same tag name
  xmlDoc ->getElementsByTagName (tagName, &nodeList);

  long len = 0;
  long i;
  nodeList ->get_length (&len);  // we're expecting to get only one.

  // The code in the following for loop will work because we will only
  // have one CustomerTypeQueryRs aggregate in the response.
  for( i = 0; i < len; ++i) {

    MSXML2::IXMLDOMNode *node = NULL;
    // Attributes Name Mapping
    MSXML2::IXMLDOMNamedNodeMap *attrNamedNodeMap = NULL;

    nodeList ->get_item (i, &node);

    // Get the attribute list
    node ->get_attributes (&attrNamedNodeMap);

    // Get the attribute values
    Utility::GetAttrValue (attrNamedNodeMap, ATTR_STATUSCODE_TAG, m_StatusCode);
    Utility::GetAttrValue (attrNamedNodeMap, ATTR_STATUSSEVERITY_TAG, m_StatusSeverity);
    Utility::GetAttrValue (attrNamedNodeMap, ATTR_STATUSMESSAGE_TAG, m_StatusMessage);

    // Clean up
    attrNamedNodeMap ->Release();
    node ->Release();
  }

  // Now walk through the DOM tree
  // to find all the CustomerTypeRet objects returned in the query.

  CComBSTR custRetTag (CUSTOMERTYPERET_TAG);

  MSXML2::IXMLDOMNodeList *custTypeNodeList = NULL;

  xmlDoc ->getElementsByTagName (custRetTag.m_str, &custTypeNodeList);

  custTypeNodeList ->get_length (&len);

  // Walk through the list of CustomerTypeRet nodes that we found
  for( i = 0; i < len; ++i) {

    // CustomerTypeRet aggregate
    MSXML2::IXMLDOMNode *node = NULL;

    MSXML2::IXMLDOMNodeList *childNodeList = NULL;
    custTypeNodeList ->get_item (i, &node);

    node ->get_childNodes (&childNodeList);
    long cIndex = 0;
    long cLen = 0;
    childNodeList ->get_length (&cLen);


    CComBSTR bstrName;
    string name;
    string value;

    QBCustomerTypeRet *tmpCustRet = new QBCustomerTypeRet();

    MSXML2::IXMLDOMNode *childNode = NULL;

    // Walk through the children of this CustomerTypeRet node,
    // using them to populate our QBCustomerTypeRet object.
    // Note that instead of iterating the children of the
    // CustomerTypeRet node, we could have asked the CustomerTypeRet
    // node for each child by name.  Then we could have
    // avoided all these strcmps!
    for (cIndex = 0; cIndex < cLen; ++cIndex) {

      childNode = NULL;
      childNodeList ->get_item (cIndex, &childNode);

      childNode ->get_nodeName (&bstrName);

      Utility::BSTRToString (bstrName, name);
      const TCHAR *szName = name.c_str();

      if (!strcmp (szName, LISTID_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->ListID (value);
      }
      else if (!strcmp (szName, TIMECREATED_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->TimeCreated (value);
      }
      else if (!strcmp (szName, TIMEMODIFIED_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->TimeModified (value);
      }
      else if (!strcmp (szName, EDITSEQUENCE_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->EditSequence (value);
      }
      else if (!strcmp (szName, NAME_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->Name (value);
      }
      else if (!strcmp (szName, FULLNAME_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->FullName (value);
      }
      else if (!strcmp (szName, ISACTIVE_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->IsActive (value);
      }
      else if (!strcmp (szName, SUBLEVEL_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->Sublevel (value);
      }

    }// done processing one CustomerTypeRet node

    // push our Ret object onto the vector
    m_CustTypeRetVec.push_back (tmpCustRet);

    // Clean up
    node ->Release();
    childNodeList ->Release();
  }

  // CustomerType Count
  m_CustomerTypeCount = i;

  // Clean up
  nodeList ->Release();
  custTypeNodeList ->Release();
  xmlDoc ->Release();

  return true;
}
