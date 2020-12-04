/*---------------------------------------------------------------------------
 * FILE: QBCustomerTypeAdd.cpp
 *
 * Description:
 * Implementation of CustomerType Add operation.
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
#include "QBCustomerTypeAdd.h"
#include "qbXMLRPWrapper.h"
#include "Utility.h"
#include "qbXMLTags.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

QBCustomerTypeAdd::QBCustomerTypeAdd(const string &companyFileName,
                                     qbXMLRPWrapper *pQbxmlrpWrapper)
  :m_CustomerTypeRet(NULL)
{
  m_CompanyFileName = companyFileName;
  m_pQbxmlrpWrapper = pQbxmlrpWrapper;
}

QBCustomerTypeAdd::~QBCustomerTypeAdd()
{

}


/*---------------------------------------------------------------------------
 * AddCustomerType
 * Add a customer type into QuickBooks.  It creates the request XML,
 * sends to QuickBooks, and parses the response XML. It uses the base class
 * method PerformOperation().
 *
 */
bool QBCustomerTypeAdd::AddCustomerType()
{
  return PerformOperation();
}


/*---------------------------------------------------------------------------
 * CreateXMLBody
 * Create CustomerTypeAddRq XML fragment.
 * You should set the data with CustomerTypeRet() before calling this method.
 *
 */
bool QBCustomerTypeAdd::CreateXMLBody()
{
  if (m_CustomerTypeRet == NULL) { // nothing to add
    m_ErrorMsg = "Nothing to add!";
    return false;
  }

  string attrsName[1];
  string attrsValue[1];

  int attrIndex = 0;
  string currRequestID;

  Utility::GetRequestID (currRequestID);
  attrsName[attrIndex] = string (ATTR_REQUESTID_TAG);
  attrsValue[attrIndex] = currRequestID;

  attrIndex++;

  m_Builder ->AddAggregate (CUSTOMERTYPEADDRQ_TAG, attrsName, attrsValue, attrIndex);
  m_Builder ->AddAggregate (CUSTOMERTYPEADD_TAG);

  m_Builder ->AddElement (NAME_TAG, m_CustomerTypeRet ->FullName());

  m_Builder ->AddEndAggregate (CUSTOMERTYPEADD_TAG);
  m_Builder ->AddEndAggregate (CUSTOMERTYPEADDRQ_TAG);

  if (m_Builder ->GetHasError()) {
    m_ErrorMsg = m_Builder ->GetErrorMsg();
    return false;
  }

  return true;
}


/*---------------------------------------------------------------------------
 * ParseXML
 * Uses the MSXML 4.0 DOM obj to parse the response XML.  ParseXML only
 * retrieves the status attributes from CustomerTypeAddRs, but you
 * can add additional functionality to walk through the child nodes
 * to get more detail if you need to.
 *
 */
bool QBCustomerTypeAdd::ParseXML()
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
  // to find the CustomerTypeAddRs Element and
  // its statusCode, statusSeverity, statusMessage attributes

  MSXML2::IXMLDOMNodeList *nodeList = NULL;

  CComBSTR tagName (CUSTOMERTYPEADDRS_TAG);

  // Get the list of elements with the same tag name
  xmlDoc ->getElementsByTagName (tagName, &nodeList);

  long i;
  long len = 0;
  nodeList ->get_length (&len);  // we're expecting to get only one.

  // The code in the following for loop will work because we will only
  // have one CustomerTypeAddRs aggregate in the response.
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

  // Clean up
  nodeList ->Release();
  xmlDoc ->Release();

  return true;
}
