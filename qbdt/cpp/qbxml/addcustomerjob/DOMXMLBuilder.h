/*---------------------------------------------------------------------------
 * FILE: DOMXMLBuilder.h
 *
 * Description:
 * Header file of DOMXMLBuilder Class. DOMXMLBuilder provides basic
 * methods to build an XML document using MSXML.
 *
 * Created On: 11/11/2001
 * Updated to SDK 2.0: 08/08/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#if !defined(AFX_DOMXMLBUILDER_H__7B602C96_B9BB_4FA6_992A_90009C2E4340__INCLUDED_)
#define AFX_DOMXMLBUILDER_H__7B602C96_B9BB_4FA6_992A_90009C2E4340__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class DOMXMLBuilder
{
public:
  DOMXMLBuilder();
  ~DOMXMLBuilder();

  // Get XML
  string & GetXML();

  // XML builder methods
  void AddAggregate(const string &inputTag,
                    const string attNameArray[] = NULL,
                    const string attValueArray[] = NULL,
                    int   attCount = 0);

  void AddEndAggregate(const string &inputTag);

  void AddElement(const string &inTag, const string &value);

  void CreateXMLTrailer();    // Create XML Trailer
  void CreateXMLHeader();     // Create XML Header

  // Error Status access methods
  bool GetHasError() { return m_HasError;};

  string &GetErrorMsg(){return m_ErrorMsg;};

private:

  // members
  bool   m_IsTopLevel;  // top level flag

  string m_XML;         // XML
  string m_Prolog;      // prolog of qbXML

  bool   m_HasError;    // Is there an error?
  string m_ErrorMsg;    // Error Message

  // DOM tree objects
  MSXML2::IXMLDOMDocument *m_XMLDoc;      // DOM document
  MSXML2::IXMLDOMNode     *m_CurrentNode; // Current Node


  void AddTagTopLevel(const TCHAR *inputStr);
  bool InstantiateXMLDomDocument();
};

#endif // !defined(AFX_DOMXMLBUILDER_H__7B602C96_B9BB_4FA6_992A_90009C2E4340__INCLUDED_)
