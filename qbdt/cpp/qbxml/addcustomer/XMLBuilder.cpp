/*---------------------------------------------------------------------------
 * FILE: XMLBuilder.cpp
 *
 * Description:
 * XMLBuilder Class Implementation. It provides the basic methods to
 * create a well formatted XML doc. It also does the XML predefined
 * chars mapping, such as map "<" to "&lt;"
 *
 * In the QuickBooks qbXML sample applications, we provide 2 classes to
 * build XML, one is XMLBuilder which is using string concatenation, the
 * other is using MSXML DOM object to build XML DOM tree. Both methods
 * has its advantage and disadvantage. The advantage of string formatted
 * XMLbuilder is it's easy, not much overhead. The disadvantage is it might
 * not be able to handle complicated application. The advantage of DOM
 * XML builder is that it is using DOM object which can handle complicated
 * application. The disadvantage is the overhead to create a DOM tree,
 * as it can be very expensive.
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
#include "XMLBuilder.h"


// constants
static const string s_indentationChars = ("    "); // 4 spaces
static const string s_endOfLine ="\r\n";   // cr and nl


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

XMLBuilder::XMLBuilder()
  : m_IndentLevel(0)
{

}

XMLBuilder::~XMLBuilder()
{

}


/*---------------------------------------------------------------------------
 * AddTagTopLevel
 * Add XML top level tags
 *
 */
void XMLBuilder::AddTagTopLevel(const string &inputStr)
{
  Append("<");
  Append(inputStr);
  Append(">");
  Append(s_endOfLine);
}


/*---------------------------------------------------------------------------
 * AddAggregate
 * Add XML Aggregate start tag.  You can specify attributes here, if any.
 *
 */
void XMLBuilder::AddAggregate(const string  &inputTag,
                              const string  attNameArray[],
                              const string  attValueArray[],
                              int           attCount)
{
  AddIndentation();

  Append("<");
  Append(inputTag);

  // add the attributes, if any
  for (int i = 0; i < attCount; ++i) {
    Append(" ");
    Append(attNameArray[i]);
    Append("=");
    Append("\"");

    // pre-defined chars handling
    // handle XML predefined chars
    if (HavePreDefinedChar(attValueArray[i])){
      // Create a replace str
      string tmpValue;
      ReplacePreDefinedChar(attValueArray[i], tmpValue);
      Append(tmpValue);
    }
    else {
      Append(attValueArray[i]);
    }
    Append("\"");
  }

  // end of tag
  Append(">");
  Append(s_endOfLine);

  m_IndentLevel++ ;
}


/*---------------------------------------------------------------------------
 * AddEndAggregate
 * Add XML Aggregate end tag
 *
 */
void XMLBuilder::AddEndAggregate(const string &inputTag)
{
  --m_IndentLevel;

  AddIndentation();

  Append("</");
  Append(inputTag);
  Append(">");
  Append(s_endOfLine);

}


/*--------------------------------------------------------------
 * AddElement
 * Add XML Element
 *
 */
void XMLBuilder::AddElement(const string &inTag,
                            const string &value)
{
  if (value.empty() ||
    value.size() == 0 ) {
    return;
  }

  AddIndentation();

  Append("<");
  Append(inTag);
  Append(">");

  // handle XML predefined chars
  if (HavePreDefinedChar(value)) {
    // Replace them
    string newValue;

    ReplacePreDefinedChar(value, newValue);
    Append(newValue);
  }
  else {
    Append(value);
  }

  Append("</");
  Append(inTag);
  Append(">");
  Append(s_endOfLine);

}


/*---------------------------------------------------------------------------
 * AddIndentation
 * Add white space for formatting
 *
 */
void XMLBuilder::AddIndentation()
{
  // make indentation
  for (int i = 1; i <= m_IndentLevel; i++) {
    m_XML += s_indentationChars;
  }
}


/*---------------------------------------------------------------------------
 * PreDefinedCharMap
 * XML predefined char map: such as map "<" to "&lt;"
 *
 */
void XMLBuilder::PreDefinedCharMap(TCHAR    inChar,
                                   string   &mappedCharStr)
{
  switch (inChar) {
    case ('>'):
      mappedCharStr = "&gt;";
      break;
    case ('<'):
      mappedCharStr = "&lt;";
      break;
    case ('&'):
      mappedCharStr = "&amp;";
      break;
    case ('\''):
      mappedCharStr = "&apos;";
      break;
    case ('"'):
      mappedCharStr = "&quot;";
      break;
    default:
      mappedCharStr = inChar;
      break;
  }

}


/*---------------------------------------------------------------------------
 * HavePreDefinedChar
 * check string to see if it contains any XML predefined char
 *
 */
bool XMLBuilder::HavePreDefinedChar(const string &inStr)
{
  if(inStr.empty()){
    return false;
  }

  int len = inStr.size();
  for (int i = 0; i < len; ++i) {
    if (IsPreDefinedChar(inStr[i])) {
      return true;
    }
  }
  return false;
}


/*---------------------------------------------------------------------------
 * IsPreDefinedChar
 * check the input char to see if it is an XML predefined char
 *
 */
bool XMLBuilder::IsPreDefinedChar(TCHAR inChar)
{
  switch (inChar) {
    case ('>'):
    case ('<'):
    case ('&'):
    case ('\''):
    case ('"'):
      return true;
    default:
      return false;

  }
}


/*---------------------------------------------------------------------------
 * ReplacePreDefinedChar
 * Replace XML predefined chars
 *
 */
void XMLBuilder::ReplacePreDefinedChar(const string   &inStr,
                                       string         &outStr)
{
  // Just return the original string if it was empty
  // or it did not contain any bad characters.
  if (inStr.empty() || !HavePreDefinedChar(inStr)){
    outStr = inStr;
    return;
  }

  outStr.erase();

  int     len = inStr.size();
  TCHAR   cChar;
  string  mappedStr;

  for (int i = 0; i < len; ++i) {
    cChar = inStr[i];
    if (IsPreDefinedChar(cChar)) { // replace the predefined char
      PreDefinedCharMap(cChar, mappedStr);
      outStr += mappedStr;
    }
    else {
      outStr += cChar;
    }
  }
}
