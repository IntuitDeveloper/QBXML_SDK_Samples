/*---------------------------------------------------------------------------
 * FILE: QBCustomerQuery.cpp
 *
 * Description:
 * Implementation of Customer Query operation.
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
#include "QBCustomerQuery.h"
#include "qbXMLRPWrapper.h"
#include "Utility.h"
#include "qbXMLTags.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

QBCustomerQuery::QBCustomerQuery(const string &companyFileName,
                                 qbXMLRPWrapper *pQbxmlrpWrapper)
  :m_CustomerCount(0)
{
  m_CompanyFileName = companyFileName;
  m_pQbxmlrpWrapper = pQbxmlrpWrapper;
}

QBCustomerQuery::~QBCustomerQuery()
{
  // clean the customer list
  QBCustRetVec::iterator it;

  for (it = m_CustRetVec.begin(); it != m_CustRetVec.end(); it++) {
    QBCustomerRet *custRet = *it;
    delete custRet;
  }
}


/*---------------------------------------------------------------------------
 * GetCustomerListFromQB
 * Get all customers from QuickBooks.  It creates the request XML,
 * sends to QuickBooks, and parses the response XML. It uses the base class
 * method PerformOperation().
 *
 */
bool QBCustomerQuery::GetCustomerListFromQB()
{
  return PerformOperation();
}


/*---------------------------------------------------------------------------
 * CreateXMLBody
 * Create CustomerQueryRq XML fragment.
 *
 */
bool QBCustomerQuery::CreateXMLBody()
{
  string attrsName[1];
  string attrsValue[1];

  int attrIndex = 0;

  string currRequestID;
  Utility::GetRequestID (currRequestID);
  attrsName[attrIndex] = ATTR_REQUESTID_TAG;
  attrsValue[attrIndex] = currRequestID;

  attrIndex++;

  // the query to get all customers
  m_Builder -> AddAggregate (CUSTOMERQUERYRQ_TAG, attrsName, attrsValue, attrIndex);
  m_Builder -> AddEndAggregate (CUSTOMERQUERYRQ_TAG);

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
bool QBCustomerQuery::ParseXML()
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
  // to find the CustomerQueryRs Element and
  // its statusCode, statusSeverity, statusMessage attributes

  MSXML2::IXMLDOMNodeList *nodeList = NULL;

  CComBSTR tagName (CUSTOMERQUERYRS_TAG);

  // Get the list of elements with the same tag name
  xmlDoc ->getElementsByTagName (tagName, &nodeList);

  long i;
  long len = 0;
  nodeList ->get_length (&len);  // we're expecting to get only one.

  // The code in the following for loop will work because we will only
  // have one CustomerQueryRs aggregate in the response.
  for(i = 0; i < len; ++i) {

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
  // to find all the CustomerRet objects returned in the query.

  CComBSTR custRetTag (CUSTOMERRET_TAG);

  MSXML2::IXMLDOMNodeList *custNodeList = NULL;

  xmlDoc ->getElementsByTagName (custRetTag.m_str, &custNodeList);

  custNodeList ->get_length (&len);

  // Walk through the list of CustomerRet nodes that we found
  for( i = 0; i < len; ++i) {

    // CustomerRet aggregate
    MSXML2::IXMLDOMNode *node = NULL;

    MSXML2::IXMLDOMNodeList *childNodeList = NULL;
    custNodeList ->get_item (i, &node);

    node ->get_childNodes (&childNodeList);
    long cIndex = 0;
    long cLen = 0;
    childNodeList ->get_length (&cLen);


    CComBSTR bstrName;
    string name;
    string value;

    QBCustomerRet *tmpCustRet = new QBCustomerRet();

    MSXML2::IXMLDOMNode *childNode = NULL;

    // Walk through the children of this CustomerRet node,
    // using them to populate our QBCustomerRet object.
    // Note that instead of iterating the children of the
    // CustomerRet node, we could have asked the CustomerRet
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
      else if (!strcmp (szName, COMPANYNAME_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->CompanyName (value);
      }
      else if (!strcmp (szName, FIRSTNAME_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->FirstName (value);
      }
      else if (!strcmp (szName, LASTNAME_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->LastName (value);
      }
      // BillAddress_* is not implemented
      else if (!strcmp (szName, PHONE_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->Phone (value);
      }
      else if (!strcmp (szName, EMAIL_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->Email (value);
      }
      else if (!strcmp (szName, CONTACT_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->Contact (value);
      }
      else if (!strcmp (szName, BALANCE_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->Balance (value);
      }
      else if (!strcmp (szName, TOTALBALANCE_TAG)) {

        Utility::GetNodeValue (childNode, value);
        tmpCustRet ->TotalBalance (value);
      }
      else if (!strcmp (szName, CUSTOMERTYPEREF_TAG)) {
        // go through the child Node to get the TypeRef
        MSXML2::IXMLDOMNodeList *typeRefNodeList = NULL;

        childNode ->get_childNodes (&typeRefNodeList);
        long refIndex = 0;
        long refLen = 0;
        typeRefNodeList ->get_length (&refLen);

        // loop through the child nodes
        MSXML2::IXMLDOMNode *refChildNode = NULL;
        for (refIndex = 0; refIndex < refLen; ++ refIndex) {
          refChildNode = NULL;
          typeRefNodeList ->get_item (refIndex, &refChildNode);

          refChildNode ->get_nodeName (&bstrName);

          Utility::BSTRToString (bstrName, name);
          const TCHAR *szName = name.c_str();

          if (!strcmp (szName, LISTID_TAG)) {

            Utility::GetNodeValue (refChildNode, value);
            tmpCustRet ->CustomerTypeRef_ListID (value);
          }
          else if (!strcmp (szName, FULLNAME_TAG)) {

            Utility::GetNodeValue (refChildNode, value);
            tmpCustRet ->CustomerTypeRef_FullName (value);
          }
        }
        typeRefNodeList ->Release();
      }
      // other Refs are not implemented here
    }// done processing one CustomerTypeRet node

    // push our Ret object onto the vector
    m_CustRetVec.push_back (tmpCustRet);

    // Clean up
    node ->Release();
    childNodeList ->Release();
  }

  // Customer Count
  m_CustomerCount = i;

  // Clean up
  nodeList ->Release();
  custNodeList ->Release();
  xmlDoc ->Release();

  return true;
}
