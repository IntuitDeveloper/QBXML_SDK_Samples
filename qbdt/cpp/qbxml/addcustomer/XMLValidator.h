/*---------------------------------------------------------------------------
 * FILE: XMLValidator.h
 *
 * Description:
 * XML Validation
 *
 * Created On: 11/07/2001
 * Updated to QBXML 2.0 August 2002
 *
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 */
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_XMLVALIDATOR_H__5C335C40_903F_4404_92E1_4883288575E9__INCLUDED_)
#define AFX_XMLVALIDATOR_H__5C335C40_903F_4404_92E1_4883288575E9__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class XMLValidator
{
public:
  XMLValidator();
  virtual ~XMLValidator();

  bool XMLValidation(const string &xmlStr,
                     const string &schemaFile);

  string GetErrorMsg() { return m_ErrorMsg;};

private:
  string m_ErrorMsg;
  void RemoveDocType(const string &xmlStr, string &newStr);
};

#endif // !defined(AFX_XMLVALIDATOR_H__5C335C40_903F_4404_92E1_4883288575E9__INCLUDED_)
