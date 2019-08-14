/*---------------------------------------------------------------------------
 * FILE: AddCustomerJob.cpp
 *
 * Description:
 * Implementation of WinMain
 * Most of this code is generated by the ATL COM AppWizard.
 * We create the dialog window in the WinMain.
 *
 * Created On: 11/11/2001
 * Updated to SDK 2.0: 08/08/2002
 *
 * Copyright � 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 */


// Note: Proxy/Stub Information
//      To build a separate proxy/stub DLL,
//      run nmake -f AddCustomerJobps.mk in the project directory.

#include "stdafx.h"
#include "resource.h"
#include <initguid.h>
#include "AddCustomerJob.h"

#include "AddCustomerJob_i.c"

#include "QBCompanyFileDlg.h"

const DWORD dwTimeOut = 5000; // time for EXE to be idle before shutting down
const DWORD dwPause = 1000; // time to wait for threads to finish up

// Passed to CreateThread to monitor the shutdown event
static DWORD WINAPI MonitorProc(void* pv)
{
  CExeModule* p = (CExeModule*)pv;
  p->MonitorShutdown();
  return 0;
}

LONG CExeModule::Unlock()
{
  LONG l = CComModule::Unlock();
  if (l == 0)
  {
    bActivity = true;
    SetEvent(hEventShutdown); // tell monitor that we transitioned to zero
  }
  return l;
}

//Monitors the shutdown event
void CExeModule::MonitorShutdown()
{
  while (1)
  {
    WaitForSingleObject(hEventShutdown, INFINITE);
    DWORD dwWait=0;
    do
    {
      bActivity = false;
      dwWait = WaitForSingleObject(hEventShutdown, dwTimeOut);
    } while (dwWait == WAIT_OBJECT_0);
    // timed out
    if (!bActivity && m_nLockCnt == 0) // if no activity let's really bail
    {
#if _WIN32_WINNT >= 0x0400 & defined(_ATL_FREE_THREADED)
      CoSuspendClassObjects();
      if (!bActivity && m_nLockCnt == 0)
#endif
        break;
    }
  }
  CloseHandle(hEventShutdown);
  PostThreadMessage(dwThreadID, WM_QUIT, 0, 0);
}

bool CExeModule::StartMonitor()
{
  hEventShutdown = CreateEvent(NULL, false, false, NULL);
  if (hEventShutdown == NULL)
    return false;
  DWORD dwThreadID;
  HANDLE h = CreateThread(NULL, 0, MonitorProc, this, 0, &dwThreadID);
  return (h != NULL);
}

CExeModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()


LPCTSTR FindOneOf(LPCTSTR p1, LPCTSTR p2)
{
  while (p1 != NULL && *p1 != NULL)
  {
    LPCTSTR p = p2;
    while (p != NULL && *p != NULL)
    {
      if (*p1 == *p)
        return CharNext(p1);
        p = CharNext(p);
      }
      p1 = CharNext(p1);
    }
  return NULL;
}

/////////////////////////////////////////////////////////////////////////////
//
extern "C" int WINAPI _tWinMain(HINSTANCE hInstance,
  HINSTANCE /*hPrevInstance*/, LPTSTR lpCmdLine, int /*nShowCmd*/)
{
  lpCmdLine = GetCommandLine(); //this line necessary for _ATL_MIN_CRT

#if _WIN32_WINNT >= 0x0400 & defined(_ATL_FREE_THREADED)
  HRESULT hRes = CoInitializeEx(NULL, COINIT_MULTITHREADED);
#else
  HRESULT hRes = CoInitialize(NULL);
#endif
  _ASSERTE(SUCCEEDED(hRes));
  _Module.Init(ObjectMap, hInstance, &LIBID_ADDCUSTOMERJOBLib);
  _Module.dwThreadID = GetCurrentThreadId();
  TCHAR szTokens[] = _T("-/");

  int nRet = 0;
  BOOL bRun = TRUE;
  LPCTSTR lpszToken = FindOneOf(lpCmdLine, szTokens);
  while (lpszToken != NULL)
  {
    if (lstrcmpi(lpszToken, _T("UnregServer"))==0)
    {
      _Module.UpdateRegistryFromResource(IDR_AddCustomerJob, FALSE);
      nRet = _Module.UnregisterServer(TRUE);
      bRun = FALSE;
      break;
    }
    if (lstrcmpi(lpszToken, _T("RegServer"))==0)
    {
      _Module.UpdateRegistryFromResource(IDR_AddCustomerJob, TRUE);
      nRet = _Module.RegisterServer(TRUE);
      bRun = FALSE;
      break;
    }
    lpszToken = FindOneOf(lpszToken, szTokens);
  }

  if (bRun)
  {
    _Module.StartMonitor();
#if _WIN32_WINNT >= 0x0400 & defined(_ATL_FREE_THREADED)
    hRes = _Module.RegisterClassObjects(CLSCTX_LOCAL_SERVER,
           REGCLS_MULTIPLEUSE | REGCLS_SUSPENDED);
           _ASSERTE(SUCCEEDED(hRes));
    hRes = CoResumeClassObjects();
#else
    hRes = _Module.RegisterClassObjects(CLSCTX_LOCAL_SERVER,
                                        REGCLS_MULTIPLEUSE);
#endif
    _ASSERTE(SUCCEEDED(hRes));

    // ----------------------------------------------------
    // Create the company file dialog window
    //
    CQBCompanyFileDlg cQBCompanyFileDlg;

    cQBCompanyFileDlg.Create(NULL);

    cQBCompanyFileDlg.ShowWindow(10);

    MSG msg;
    while (GetMessage(&msg, 0, 0, 0)) {
      TranslateMessage(&msg);
      DispatchMessage(&msg);
    }

     _Module.RevokeClassObjects();
     Sleep(dwPause); //wait for any threads to finish
  }

  _Module.Term();
  CoUninitialize();
  return nRet;
}