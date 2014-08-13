/*---------------------------------------------------------------------------
 * FILE: QBSDKCallback.h
 *
 * Description:
 * Header file of CQBSDKCallback. 
 * Defines callback for QB event notification.
 *
 * Created On: 09/15/2003
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#if !defined(AFX_QBSDKCALLBACK_H__5AE97D5E_ABB1_4B8A_9DBA_4B8B46308ADE__INCLUDED_)
#define AFX_QBSDKCALLBACK_H__5AE97D5E_ABB1_4B8A_9DBA_4B8B46308ADE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"       // main symbols

#import  "sdkevent.dll" no_namespace named_guids


/////////////////////////////////////////////////////////////////////////////
// CQBSDKCallback

class ATL_NO_VTABLE CQBSDKCallback : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CQBSDKCallback,&CLSID_QBSDKCallback>,
    IQBEventCallback
{
public:
	CQBSDKCallback() {}
BEGIN_COM_MAP(CQBSDKCallback)
	//COM_INTERFACE_ENTRY(IQBSDKCallback)
	COM_INTERFACE_ENTRY(IQBEventCallback)
END_COM_MAP()
DECLARE_PROTECT_FINAL_CONSTRUCT()

DECLARE_REGISTRY_RESOURCEID(IDR_QBSDKCallback)

// IQBSDKCallback
public:
     STDMETHOD(inform)(/*[in]*/ BSTR eventXML);
     HRESULT __stdcall raw_inform (/*[in]*/ BSTR eventXML )  { return inform(eventXML) ;};

private:
    void   WriteToFile( LPCTSTR lpszPath, LPCTSTR lpszXML );
    void   RecoveryAlert( const string& strCompanyFile );
    string GetTagValue( const string& strXML, const string& strTag );
};

#endif // !defined(AFX_QBSDKCALLBACK_H__5AE97D5E_ABB1_4B8A_9DBA_4B8B46308ADE__INCLUDED_)
