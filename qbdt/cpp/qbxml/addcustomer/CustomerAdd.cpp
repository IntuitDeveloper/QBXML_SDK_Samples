/*---------------------------------------------------------------------------
 * FILE: CustomerAdd.cpp
 *
 * Description:
 * Implementation of CustomerAdd
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
#include "CustomerAdd.h"
#include "qbXMLWrapper.h"
#include "XMLValidator.h"
#include "Utility.h"

#include "qbXMLTags.h"


// Constructor
//
CCustomerAdd::CCustomerAdd(const string &companyFileName,
                           const string &custName,
                           const string &firstName,
                           const string &lastName,
                           const string &phone)
  : m_CompanyFileName(companyFileName),
    m_CustomerName(custName),
    m_FirstName(firstName),
    m_LastName(lastName),
    m_Phone(phone)
{
  // Set Error Status to false, no error yet
  m_ErrorStatus = false;

  m_Msg = "OK";

  // Create XML Builder
  m_Builder = new XMLBuilder();
}


// Destructor
//
CCustomerAdd::~CCustomerAdd()
{
  delete m_Builder;
}


/*---------------------------------------------------------------------------
 * BuildXML
 * Build the CustomerAdd request XML using XMLBuilder
 *
 */
void CCustomerAdd::BuildXML()
{

  m_RequestXML = "";

  // Create QBXML Header
  CreateXMLHeader();

  // Create QBXML CustomerAddRq Block
  CreateCustomerAddRq();

  // Create QBXML trailer
  CreateXMLTrailer();

  // get the XML
  m_RequestXML = m_Builder ->GetXML();
  if (m_RequestXML.empty() ||
      m_RequestXML == "") {
    m_ErrorStatus = true;
    m_Msg = "Internal error with XMLBuilder!";
    return;
  }
}


/*---------------------------------------------------------------------------
 * DoRequest
 * Use the qbXMLRPWrapper to connect to qbXML COM, send the
 * requestXML, and get the response XML
 *
 * We assume that you have already called BuildXML() prior
 * to calling this method.
 *
 */
void CCustomerAdd::DoRequest()
{

  // Create a new qbXMLRPWrapper obj
  qbXMLRPWrapper qbXMLRPW(m_CompanyFileName,
                          m_RequestXML);

  // connect to qbXML COM and get the response
  m_ResponseXML = qbXMLRPW.QBXMLRPCall();

  m_ErrorStatus = qbXMLRPW.GetErrorStatus();
  m_Msg = qbXMLRPW.GetMsg();

}


/*---------------------------------------------------------------------------
 * ParseResponseXML
 * Use MSXML 4.0 DOM obj to parse the response XML. ParseResponseXML only
 * retrieves statusCode, statusSeverity, and statusMessage from CustomerAddRs,
 * but it is easy to add additional functionality later to walk through the
 * child nodes to get more detail.
 *
 * We assume that you've already called DoRequest() prior
 * to calling this method.
 *
 */
void CCustomerAdd::ParseResponseXML()
{

  // Initialize COM
  //if (FAILED(CoInitialize(NULL))) {
  //  m_Msg = "Unable to initialize COM";
  //  m_ErrorStatus = true;
  //  return;
  //}

  // Create xmlDoc Obj
  MSXML2::IXMLDOMDocument *xmlDoc = NULL;

  string reason;

  // Load XML
  xmlDoc = Utility::LoadXML(m_ResponseXML, reason);

  if (xmlDoc == NULL){
    m_Msg = reason;
    m_ErrorStatus = true;
    return;
  }

  // walk through the DOM Tree
  // to find the CustomerAddRs Element and
  // its statusCode, statusSeverity, statusMessage attributes

  MSXML2::IXMLDOMNodeList *nodeList = NULL;

  CComBSTR tagName (CUSTOMERADDRS_TAG);

  // Get the elements list with the same tag name
  xmlDoc ->getElementsByTagName(tagName.m_str, &nodeList);

  long len = 0;
  nodeList ->get_length(&len);  // we're expecting to get only one.

  // The code in the following for loop will work because we will only
  // have one CustomerAddRs aggregate in the response.
  for (long i = 0; i < len; ++i) {

    MSXML2::IXMLDOMNode *node = NULL;
    // Attributes Name Mapping
    MSXML2::IXMLDOMNamedNodeMap *attrNamedNodeMap = NULL;

    nodeList ->get_item(i, &node);

    // Get the attribute list
    node ->get_attributes(&attrNamedNodeMap);

    // Get the attribute values
    Utility::GetAttrValue(attrNamedNodeMap, ATTR_STATUSCODE_TAG, m_StatusCode);
    Utility::GetAttrValue(attrNamedNodeMap, ATTR_STATUSSEVERITY_TAG, m_StatusSeverity);
    Utility::GetAttrValue(attrNamedNodeMap, ATTR_STATUSMESSAGE_TAG, m_StatusMessage);

    // Clean up
    attrNamedNodeMap ->Release();
    node ->Release();
  }

  // Clean up
  nodeList ->Release();
  xmlDoc ->Release();
  //CoUninitialize();
}


/*---------------------------------------------------------------------------
 * CreateXMLHeader
 * Private method, to create Top Level of the request XML
 *
 */
void CCustomerAdd::CreateXMLHeader()
{

  // Top Level
  m_Builder -> AddTagTopLevel ("?xml version=\"1.0\" ?");

  m_Builder -> AddTagTopLevel ("?qbxml version=\"2.0\" ?");

  // QBXML starts here
  m_Builder -> AddAggregate (QBXML_TAG);

  string attrsName[1];
  string attrsValue[1];

  int attrIndex = 0;

  attrsName[attrIndex] = string(ATTR_ONERROR_TAG);
  attrsValue[attrIndex] = "stopOnError";
  attrIndex++;

  m_Builder -> AddAggregate (string(QBXMLMSGSRQ_TAG), attrsName, attrsValue, attrIndex);
}


/*---------------------------------------------------------------------------
 * CreateXMLTrailer
 * Private method, to create tailer of the request XML
 *
 */
void CCustomerAdd::CreateXMLTrailer()
{
  m_Builder -> AddEndAggregate (QBXMLMSGSRQ_TAG);
  m_Builder -> AddEndAggregate (QBXML_TAG);
}


/*---------------------------------------------------------------------------
 * CreateCustomerAddRq
 * Private method, to create QBXML's CustomerAddRq Block
 *
 */
void CCustomerAdd::CreateCustomerAddRq()
{

  string attrsName[1];
  string attrsValue[1];

  int attrIndex = 0;

  string rID;
  Utility::GetRequestID(rID);
  attrsName[attrIndex] = string(ATTR_REQUESTID_TAG);
  attrsValue[attrIndex] = rID;

  attrIndex++;
  m_Builder -> AddAggregate (string(CUSTOMERADDRQ_TAG), attrsName, attrsValue, attrIndex);
  m_Builder -> AddAggregate (CUSTOMERADD_TAG);

  m_Builder -> AddElement (NAME_TAG, m_CustomerName);
  m_Builder -> AddElement (FIRSTNAME_TAG, m_FirstName);
  m_Builder -> AddElement (LASTNAME_TAG, m_LastName);
  m_Builder -> AddElement (PHONE_TAG, m_Phone);

  m_Builder -> AddEndAggregate (CUSTOMERADD_TAG);
  m_Builder -> AddEndAggregate (CUSTOMERADDRQ_TAG);
}

