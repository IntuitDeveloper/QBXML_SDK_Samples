/*---------------------------------------------------------------------------
 * FILE: QBCustomerAdd.cpp
 *
 * Description:
 * Implementation of Customer Add operation.
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
#include "QBCustomerAdd.h"
#include "qbXMLRPWrapper.h"
#include "Utility.h"
#include "qbXMLTags.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

QBCustomerAdd::QBCustomerAdd(const string &companyFileName,
                             qbXMLRPWrapper *pQbxmlrpWrapper)
  :m_CustomerRet(NULL)
{
  m_CompanyFileName = companyFileName;
  m_pQbxmlrpWrapper = pQbxmlrpWrapper;
}

QBCustomerAdd::~QBCustomerAdd()
{

}


/*---------------------------------------------------------------------------
 * AddCustomer
 * Add a customer into QuickBooks.  It creates the request XML,
 * sends to QuickBooks, and parses the response XML. It uses the base class
 * method PerformOperation().
 *
 */
bool QBCustomerAdd::AddCustomer()
{
  return PerformOperation();
}


/*---------------------------------------------------------------------------
 * CreateXMLBody
 * Create CustomerAddRq XML fragment.
 * You should set the data with CustomerRet() before calling this method.
 *
 */
bool QBCustomerAdd::CreateXMLBody()
{
  if (m_CustomerRet == NULL) { // nothing to add
    m_ErrorMsg = "Nothing to add!";
    return false;
  }

  string attrsName[1];
  string attrsValue[1];

  int attrIndex = 0;

  string currRequestID;
  Utility::GetRequestID (currRequestID);
  attrsName[attrIndex] = ATTR_REQUESTID_TAG;
  attrsValue[attrIndex] = currRequestID;

  attrIndex++;

  m_Builder ->AddAggregate (CUSTOMERADDRQ_TAG, attrsName, attrsValue, attrIndex);

  m_Builder ->AddAggregate (CUSTOMERADD_TAG);

  m_Builder ->AddElement (NAME_TAG, m_CustomerRet ->Name());

  // Parent ref
  m_Builder ->AddAggregate (PARENTREF_TAG);
  m_Builder ->AddElement (FULLNAME_TAG, m_CustomerRet ->ParentRef_FullName());
  m_Builder ->AddEndAggregate (PARENTREF_TAG);

  m_Builder ->AddElement (COMPANYNAME_TAG, m_CustomerRet ->CompanyName());
  m_Builder ->AddElement (PHONE_TAG, m_CustomerRet ->Phone());
  m_Builder ->AddElement (EMAIL_TAG, m_CustomerRet ->Email());

  // CustomerTypeRef
  if (m_CustomerRet->CustomerTypeRef_FullName() != "")
  {
	  m_Builder ->AddAggregate (CUSTOMERTYPEREF_TAG);
	  m_Builder ->AddElement (FULLNAME_TAG, m_CustomerRet ->CustomerTypeRef_FullName());
	  m_Builder ->AddEndAggregate (CUSTOMERTYPEREF_TAG);
  }

  // end aggregate
  m_Builder ->AddEndAggregate (CUSTOMERADD_TAG);
  m_Builder ->AddEndAggregate (CUSTOMERADDRQ_TAG);

  if (m_Builder ->GetHasError()) {
    m_ErrorMsg = m_Builder ->GetErrorMsg();
    return false;
  }

  return true;
}


/*---------------------------------------------------------------------------
 * ParseXML
 * Uses the MSXML 4.0 DOM obj to parse the response XML.  ParseXML only
 * retrieves the status attributes from CustomerAddRs, but you
 * can add additional functionality to walk through the child nodes
 * to get more detail if you need to.
 *
 */
bool QBCustomerAdd::ParseXML()
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
  // to find the CustomerAddRs Element and
  // its statusCode, statusSeverity, statusMessage attributes

  MSXML2::IXMLDOMNodeList *nodeList = NULL;

  CComBSTR tagName (CUSTOMERADDRS_TAG);

  // Get the list of elements with the same tag name
  xmlDoc ->getElementsByTagName (tagName, &nodeList);

  long i;
  long len = 0;
  nodeList ->get_length (&len);  // we're expecting to get only one.

  // The code in the following for loop will work because we will only
  // have one CustomerAddRs aggregate in the response.
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
