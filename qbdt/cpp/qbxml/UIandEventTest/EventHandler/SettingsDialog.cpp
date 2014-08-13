/*---------------------------------------------------------------------------
 * FILE: SettingsDialog.cpp
 *
 * Description:
 * MFC-generated dialog class.
 *
 * Created On: 09/15/2003
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#include "stdafx.h"
#include "EventHandler.h"
#include "SettingsDialog.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CSettingsDialog dialog


// MFC-added contructor.
CSettingsDialog::CSettingsDialog(CWnd* pParent /*=NULL*/)
	: CDialog(CSettingsDialog::IDD, pParent)
{
    // MFC-managed elements.
	//{{AFX_DATA_INIT(CSettingsDialog)
	m_strOutputPath = _T("");
    m_nLaunch = IDC_LAUNCH_YES;
    m_nRecoveryAlert = 0;
	//}}AFX_DATA_INIT
}


// MFC-managed data exchange
void CSettingsDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CSettingsDialog)
	DDX_Text(pDX, IDC_OUTPUT_PATH, m_strOutputPath);
	DDV_MaxChars(pDX, m_strOutputPath, MAX_PATH);
    DDX_Radio( pDX, IDC_LAUNCH_NO, m_nLaunch );
    DDX_Check( pDX, IDC_RECOVERY_ALERT, m_nRecoveryAlert );
	//}}AFX_DATA_MAP
}


// MFC-managed message map.
BEGIN_MESSAGE_MAP(CSettingsDialog, CDialog)
	//{{AFX_MSG_MAP(CSettingsDialog)
	ON_BN_CLICKED(IDC_BROWSE, OnBrowse)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSettingsDialog message handlers

// OnInitDialog()
// Dialog initialization.
BOOL CSettingsDialog::OnInitDialog() 
{
	CDialog::OnInitDialog();

    // Load current settings from registry.
    CEventHandlerApp *pApp = (CEventHandlerApp*) AfxGetApp();
    if( pApp ) {
        m_nLaunch = pApp->GetProfileInt(    pApp->lpszProfileSection,
                                            pApp->lpszProfileLaunch,
                                            1 );
        m_nRecoveryAlert = pApp->GetProfileInt( pApp->lpszProfileSection, 
                                                pApp->lpszProfileRecoveryAlert,
                                                1 );
        m_strOutputPath = pApp->GetProfileString(   pApp->lpszProfileSection, 
                                                    pApp->lpszProfileOutputPath,
                                                    "" );
        // Update controls.
        UpdateData( false );
    }
    
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

// OnOK
// Save settings when user clicks "OK" button.
void CSettingsDialog::OnOK()
{
    // Load member variables from controls.
    UpdateData( true );

    // Save settings to registry.
    CEventHandlerApp *pApp = (CEventHandlerApp*) AfxGetApp();
    if( pApp ) {
        pApp->WriteProfileInt(  pApp->lpszProfileSection,
                                pApp->lpszProfileLaunch,
                                m_nLaunch );
        pApp->WriteProfileInt(  pApp->lpszProfileSection,
                                pApp->lpszProfileRecoveryAlert,
                                m_nRecoveryAlert );
        pApp->WriteProfileString(   pApp->lpszProfileSection,
                                    pApp->lpszProfileOutputPath,
                                    m_strOutputPath );
    }
	CDialog::OnOK();
}

// Callback for OnBrowse() below.
INT CALLBACK BrowseCallbackProc( HWND    hWnd, 
                                 UINT    uMsg,
                                 LPARAM  lParam, 
                                 LPARAM  lpData) 
{
    // Set initial directory to value passed in via lpData.
    if( uMsg == BFFM_INITIALIZED ) {
        if( lpData ) {
            SendMessage( hWnd, BFFM_SETSELECTION, TRUE, lpData );
        }
    }

    return 0;
}

// OnBrowse()
// Built-in Windows browse dialog for selecting a directory.
void CSettingsDialog::OnBrowse() 
{
    // Set initial directory to value in edit box.
    TCHAR initialPath[MAX_PATH+1];
    GetDlgItemText( IDC_OUTPUT_PATH, initialPath, MAX_PATH );
    // If edit box is empty, set initial directory to current directory.
    if( _tcslen(initialPath) == 0 ) {
        GetCurrentDirectory( MAX_PATH, initialPath );
    }

    // Setup browse dialog.
    TCHAR displayName[MAX_PATH+1];
    BROWSEINFO browseInfo;
    browseInfo.hwndOwner = m_hWnd;
    browseInfo.pidlRoot = NULL;
    browseInfo.pszDisplayName = displayName;
    browseInfo.lpszTitle = _T("Select Output Directory");
    browseInfo.ulFlags = BIF_EDITBOX;
    browseInfo.lpfn = BrowseCallbackProc;
    browseInfo.lParam = (LPARAM)initialPath;
    browseInfo.iImage = 0;

    // Show browse dialog.
    LPITEMIDLIST idList = SHBrowseForFolder( &browseInfo );

    // If directory selected, write it to edit control.
    if( idList != NULL ) {
        if( SHGetPathFromIDList(idList,initialPath) ) {
            SetDlgItemText( IDC_OUTPUT_PATH, initialPath );
        }
    }
}

