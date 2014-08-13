; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CShowXMLDlg
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "menusubscribe.h"
LastPage=0

ClassCount=3
Class1=CMenuSubscribeApp
Class2=CMenuSubscribeDlg


ResourceCount=2
Resource1=IDD_SHOWXML
Class3=CShowXMLDlg
Resource2=IDD_MENUSUBSCRIBE_DIALOG

[CLS:CMenuSubscribeApp]
Type=0
HeaderFile=MenuSubscribe.h
ImplementationFile=MenuSubscribe.cpp
Filter=N
LastObject=CMenuSubscribeApp
BaseClass=CWinApp
VirtualFilter=AC

[DLG:IDD_MENUSUBSCRIBE_DIALOG]
Type=1
Class=CMenuSubscribeDlg
ControlCount=28
Control1=IDC_STATIC,static,1342308352
Control2=IDC_STATIC,static,1342308352
Control3=IDC_ADD_TO,combobox,1344339971
Control4=IDC_STATIC,static,1342308352
Control5=IDC_STATIC,button,1342177287
Control6=IDC_STATIC,button,1342177287
Control7=IDC_STATIC,button,1342177287
Control8=IDC_STATIC,static,1342308352
Control9=IDC_MENU1_TEXT,edit,1350631552
Control10=IDC_MENU1_VISIBLE_IF,button,1342246915
Control11=IDC_MENU1_VISIBLE,combobox,1344339971
Control12=IDC_MENU1_ENABLED_IF,button,1342246915
Control13=IDC_MENU1_ENABLED,combobox,1344339971
Control14=IDC_STATIC,static,1342308352
Control15=IDC_MENU2_TEXT,edit,1350631552
Control16=IDC_MENU2_VISIBLE_IF,button,1342246915
Control17=IDC_MENU2_VISIBLE,combobox,1344339971
Control18=IDC_MENU2_ENABLED_IF,button,1342246915
Control19=IDC_MENU2_ENABLED,combobox,1344339971
Control20=IDC_STATIC,static,1342308352
Control21=IDC_MENU3_TEXT,edit,1350631552
Control22=IDC_MENU3_VISIBLE_IF,button,1342246915
Control23=IDC_MENU3_VISIBLE,combobox,1344339971
Control24=IDC_MENU3_ENABLED_IF,button,1342246915
Control25=IDC_MENU3_ENABLED,combobox,1344339971
Control26=IDC_ADD_SUBSCRIPTION,button,1342242817
Control27=IDCANCEL,button,1342242816
Control28=IDC_SHOW_XML,button,1342242819

[CLS:CSettingsDialog]
Type=0
HeaderFile=MenuSubscribeDlg.h
ImplementationFile=MenuSubscribeDlg.cpp
BaseClass=CDialog
Filter=D
LastObject=IDC_SHOW_XML
VirtualFilter=dWC

[CLS:CMenuSubscribeDlg]
Type=0
HeaderFile=menusubscribedlg.h
ImplementationFile=menusubscribedlg.cpp
BaseClass=CDialog
LastObject=CMenuSubscribeDlg

[DLG:IDD_SHOWXML]
Type=1
Class=CShowXMLDlg
ControlCount=3
Control1=IDOK,button,1342242817
Control2=IDC_SHOWXML_EDIT,edit,1353779396
Control3=IDC_SHOWXML_HEADER,static,1342308352

[CLS:CShowXMLDlg]
Type=0
HeaderFile=ShowXMLDlg.h
ImplementationFile=ShowXMLDlg.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=CShowXMLDlg

