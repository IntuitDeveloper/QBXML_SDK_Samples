/*---------------------------------------------------------------------------
 * FILE: XMLBuilder.h
 *
 * Description:
 * Header file of XMLBuilder Class.
 *
 * Created On: 11/07/2001
 * Updated to QBXML 2.0 August 2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */
#if !defined(AFX_XMLBUILDER_H__1A40977C_2591_441C_81D1_FA853488EF59__INCLUDED_)
#define AFX_XMLBUILDER_H__1A40977C_2591_441C_81D1_FA853488EF59__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class XMLBuilder
{
public:

  // constructor & desctructor
  //
  XMLBuilder();
  virtual ~XMLBuilder();

  // Get XML
  string &GetXML() { return m_XML;};

  // XML builder methods
  //
  void AddTagTopLevel(const string &inputStr);

  void AddAggregate(const string &inputTag,
                    const string attNameArray[] = NULL,
                    const string attValueArray[] = NULL,
                    int    attCount = 0);

  void AddEndAggregate(const string &inputTag);

  void AddElement(const string &inTag, const string &value);

private:

  int  m_IndentLevel;    // amount of indentation for the next line

  string m_XML;          // the XML data itself

  // private methods
  //
  void AddIndentation();

  // helper methods for PreDefined Char mapping
  void PreDefinedCharMap(TCHAR inChar, string &mappedStr);
  bool HavePreDefinedChar(const string &inChar);
  bool IsPreDefinedChar(TCHAR inChar);
  void ReplacePreDefinedChar(const string &inStr, string &outStr);

  void Append(const string &value){m_XML += value;};
};

#endif // !defined(AFX_XMLBUILDER_H__1A40977C_2591_441C_81D1_FA853488EF59__INCLUDED_)
