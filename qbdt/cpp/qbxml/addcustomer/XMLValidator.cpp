/*---------------------------------------------------------------------------
 * FILE: XMLValidator.cpp
 *
 * Description:
 * XMLValidator Class Implementation. It provides a method to validate XML.
 *
 * Created On: 11/07/2001
 * Updated to QBXML 2.0 August 2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */
#include "stdafx.h"
#include "XMLValidator.h"

#import "msxml4.dll"
using namespace MSXML2;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

XMLValidator::XMLValidator()
{
  m_ErrorMsg = "Validating the XML";
}

XMLValidator::~XMLValidator()
{

}

/*---------------------------------------------------------------------------
 * XMLValidation
 * Use MSXML to validate the XML.
 * Since our DOCTYPE format is not standard XML, MSXML can't process qbXML.
 * We must remove the DOCTYPE line from qbXML request before running
 * the validation.
 *
 */
bool XMLValidator::XMLValidation(const string &inputXMLStr,
                                 const string &schemaFile)
{

  HRESULT hr;
  bool    bSuccess= true;

  if (inputXMLStr.empty() ||
      schemaFile.empty()) {
    m_ErrorMsg = "Either the XML string or the schema file name is NULL!";
    return false;
  }

  // remove the DocType
  string xmlStr;
  RemoveDocType(inputXMLStr, xmlStr);

  if (xmlStr.empty()) {
    m_ErrorMsg = "xmlStr is empty";
    return false;
  }

  // Initialize COM
  //if (FAILED(CoInitialize(NULL))) {
  //  m_ErrorMsg = "Unable to initialize COM";
  //  return false;
  //}

  // DOM Document, schema, schema cache objects
  MSXML2::IXMLDOMDocument2Ptr pXMLDoc = NULL;
  MSXML2::IXMLDOMDocument2Ptr pSchema = NULL;
  MSXML2::IXMLDOMSchemaCollection2Ptr pSchemaCache = NULL;

  try {
    _variant_t varOut((bool)TRUE);

    // Create instances
    //
    pXMLDoc.CreateInstance(__uuidof(MSXML2::DOMDocument40));
    pSchema.CreateInstance(__uuidof(MSXML2::DOMDocument40));
    pSchemaCache.CreateInstance(__uuidof(MSXML2::XMLSchemaCache40));

    // make sure our objects were created - if not, this is a config issue
    if ((pXMLDoc == NULL) ||
        (pSchema == NULL) ||
        (pSchemaCache == NULL)) {
      m_ErrorMsg = "Failed to create interfaces for MSXML. Make sure MSXML 4.0 is installed";

      return false;
    }

    // first to load schema
    hr = pSchema->put_async (VARIANT_FALSE);
    if (SUCCEEDED(hr)) {
      varOut = pSchema->load(schemaFile.c_str());
      if ((bool)varOut != TRUE) {
        IXMLDOMParseErrorPtr errPtr =  pSchema->GetparseError();
        _bstr_t bstrErr(errPtr->reason);
        m_ErrorMsg = "Failed to load Schema file or one of the schemas it includes";
        char buf[256];
        sprintf(buf, "Error, line %ld, reason: %s", errPtr->line,(char*)bstrErr);
        m_ErrorMsg += "\n";
        m_ErrorMsg += buf;
        return false;
      } else {
        if (pSchema != NULL)  {
          hr = pSchemaCache->raw_add(_bstr_t(""), _variant_t((IUnknown *)pSchema));
          if (FAILED(hr)) {
            m_ErrorMsg = "Failed to load schema cache.  Perhaps you chose the wrong schema file?";
            return false;
          }
        }
      }
    }
    if (SUCCEEDED(hr) && pXMLDoc != NULL) {
      hr = pXMLDoc->put_async(VARIANT_FALSE);
      if (pSchemaCache != NULL)  {
        // refer to this schema
        hr = pXMLDoc->putref_schemas(_variant_t((IUnknown *)pSchemaCache));
      }

      if (SUCCEEDED(hr)) {
        IXMLDOMParseErrorPtr errPtr;

        _bstr_t bs(xmlStr.c_str());
        varOut = pXMLDoc->loadXML(bs);

        if ((bool) varOut == TRUE) {
          bSuccess = true;

          // Validate the XML doc
          //
          hr = pXMLDoc->raw_validate(&errPtr);
          if (SUCCEEDED(hr)) {
            m_ErrorMsg = "This is a Valid XML";
          } else {
            bSuccess = false;
          }
        } else {
          m_ErrorMsg = "Failed to parse XML";
          bSuccess   = false;
          errPtr     = pXMLDoc->GetparseError();
        }

        if (!bSuccess)  {
          _bstr_t bstrErr(errPtr->reason);
          m_ErrorMsg = "Validation failed";
          char buf[256];
          sprintf(buf, "Error, line %ld, reason: %s", errPtr->line,(char*)bstrErr);
          m_ErrorMsg += "\n";
          m_ErrorMsg += buf;
        }
      }
    }
  }
  catch(...)
  {
    m_ErrorMsg = "Unexpected error validating response";
  }

  // Clean up
  //
  pSchemaCache.Release();
  pSchema.Release();
  pXMLDoc.Release();
  //CoUninitialize();
  return(bSuccess);

}

/*---------------------------------------------------------------------------
 * RemoveDocType
 * to remove the DOCTYPE line from XML string
 *
 */
void XMLValidator::RemoveDocType(const string &xmlStr,
                                 string       &newXMLStr)
{
  if (xmlStr.empty()) {
    return;
  }

  int position;

  // find the beginning of the qbxml version line.
  position = xmlStr.find("<?qbxml version", 0);

  if (position < 0) { // can't find the "DOCTYPE"
    return;
  }

  // get everything up to the start of the DOCTYPE line
  newXMLStr = xmlStr.substr(0, position);

  // find the end of the DOCTYPE line.
  int secondPos = xmlStr.find("\n", position);

  // append everything after the DOCTYPE line.
  newXMLStr += xmlStr.substr(secondPos + 1);
}
