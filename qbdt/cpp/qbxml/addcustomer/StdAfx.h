// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently
//


#if !defined(AFX_STDAFX_H__DA82C9B1_14EB_46F1_B677_7E5C4CF2093A__INCLUDED_)
#define AFX_STDAFX_H__DA82C9B1_14EB_46F1_B677_7E5C4CF2093A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0A00
#endif
#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
class CExeModule : public CComModule
{
public:
  LONG Unlock();
  DWORD dwThreadID;
  HANDLE hEventShutdown;
  void MonitorShutdown();
  bool StartMonitor();
  bool bActivity;
};
extern CExeModule _Module;
#include <atlcom.h>
#include <atlwin.h>


// STL lib
#include <string> // STL string

using namespace std;

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__DA82C9B1_14EB_46F1_B677_7E5C4CF2093A__INCLUDED)
