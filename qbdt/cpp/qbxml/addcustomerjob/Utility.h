/*---------------------------------------------------------------------------
 * FILE: Utility.h
 *
 * Description:
 * provides static util methods. they are:
 *
 *  1. GetAttrValue: MSXML util function, to get attribute value
 *  2. GetNodeValue: MSXML util function, to get node value
 *  3. GetRequestID: use rand() to generate a request ID
 *  4. LoadXML: LoadXML with error check.
 *
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

#if !defined(AFX_Utility_H__CA6E5D38_43D2_4D35_9F69_909DB6963B22__INCLUDED_)
#define AFX_Utility_H__CA6E5D38_43D2_4D35_9F69_909DB6963B22__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class Utility
{
public:
  Utility();
  ~Utility();

  // MSXML DOM utilities
  static HRESULT GetAttrValue(MSXML2::IXMLDOMNamedNodeMap *nodeMap,
                              const string &attrName,
                              string &attrValue);
  static HRESULT GetNodeValue(MSXML2::IXMLDOMNode *node,
                              string &nodeValue);

  static HRESULT GetNodeValue(MSXML2::IXMLDOMNode *node,
                              TCHAR *nodeValue);

  static MSXML2::IXMLDOMDocument *LoadXML(const string &xmlStr, string &reason);

  // BSTR to string
  static void BSTRToString(BSTR inVal, string &outVal);

  // generate a unique requestID string
  static void GetRequestID(string &rID);

private:
  static TCHAR * GetRequestID();

};

#endif // !defined(AFX_Utility_H__CA6E5D38_43D2_4D35_9F69_909DB6963B22__INCLUDED_)
