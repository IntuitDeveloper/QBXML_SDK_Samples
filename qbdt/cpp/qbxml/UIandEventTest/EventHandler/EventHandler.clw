; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CAboutDialog
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "EventHandler.h"
LastPage=0

ClassCount=6
Class1=CEventHandlerApp
Class3=CMainFrame
Class4=CAboutDlg

ResourceCount=4
Resource1=IDR_MAINFRAME
Resource2=IDD_ABOUTBOX
Class2=CChildView
Resource3=IDD_MAIN_FORM
Class5=CSettingsDialog
Class6=CAboutDialog
Resource4=IDD_SETTINGS

[CLS:CEventHandlerApp]
Type=0
HeaderFile=EventHandler.h
ImplementationFile=EventHandler.cpp
Filter=N
LastObject=CEventHandlerApp
BaseClass=CWinApp
VirtualFilter=AC

[CLS:CChildView]
Type=0
HeaderFile=ChildView.h
ImplementationFile=ChildView.cpp
Filter=W
BaseClass=CFormView 
VirtualFilter=VWC
LastObject=CChildView

[CLS:CMainFrame]
Type=0
HeaderFile=MainFrm.h
ImplementationFile=MainFrm.cpp
Filter=T
LastObject=ID_APP_ABOUT
BaseClass=CFrameWnd
VirtualFilter=fWC




[CLS:CAboutDlg]
Type=0
HeaderFile=EventHandler.cpp
ImplementationFile=EventHandler.cpp
Filter=D
LastObject=CAboutDlg

[DLG:IDD_ABOUTBOX]
Type=1
Class=CAboutDlg
ControlCount=6
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308480
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889
Control5=IDC_STATIC,static,1342308352
Control6=IDC_STATIC,static,1342308352

[MNU:IDR_MAINFRAME]
Type=1
Class=CMainFrame
Command1=ID_FILE_REGISTER
Command2=ID_FILE_UNREGISTER
Command3=ID_APP_EXIT
Command4=ID_EDIT_COPY
Command5=ID_EDIT_SETTINGS
Command6=ID_VIEW_STATUS_BAR
Command7=ID_APP_ABOUT
CommandCount=7

[ACL:IDR_MAINFRAME]
Type=1
Class=CMainFrame
Command1=ID_EDIT_COPY
Command2=ID_EDIT_PASTE
Command3=ID_EDIT_UNDO
Command4=ID_EDIT_CUT
Command5=ID_NEXT_PANE
Command6=ID_PREV_PANE
Command7=ID_EDIT_COPY
Command8=ID_EDIT_PASTE
Command9=ID_EDIT_CUT
Command10=ID_EDIT_UNDO
CommandCount=10

[DLG:IDD_MAIN_FORM]
Type=1
Class=CAboutDialog
ControlCount=2
Control1=IDC_HEADER,static,1342308352
Control2=IDC_EVENT_XML,edit,1353779652

[DLG:IDD_SETTINGS]
Type=1
Class=CSettingsDialog
ControlCount=8
Control1=IDC_STATIC,static,1342308352
Control2=IDC_OUTPUT_PATH,edit,1350631552
Control3=IDC_STATIC,button,1342308359
Control4=IDC_LAUNCH_NO,button,1342308361
Control5=IDC_LAUNCH_YES,button,1342177289
Control6=IDOK,button,1342373889
Control7=IDCANCEL,button,1342373888
Control8=IDC_BROWSE,button,1342242816

[CLS:CSettingsDialog]
Type=0
HeaderFile=SettingsDialog.h
ImplementationFile=SettingsDialog.cpp
BaseClass=CDialog
Filter=D
LastObject=IDC_BROWSE
VirtualFilter=dWC

[CLS:CAboutDialog]
Type=0
HeaderFile=AboutDialog.h
ImplementationFile=AboutDialog.cpp
BaseClass=CDialog
Filter=D

