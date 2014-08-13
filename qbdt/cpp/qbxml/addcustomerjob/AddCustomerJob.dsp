# Microsoft Developer Studio Project File - Name="AddCustomerJob" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Application" 0x0101

CFG=AddCustomerJob - Win32 Release
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "AddCustomerJob.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "AddCustomerJob.mak" CFG="AddCustomerJob - Win32 Release"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "AddCustomerJob - Win32 Debug" (based on "Win32 (x86) Application")
!MESSAGE "AddCustomerJob - Win32 Release" (based on "Win32 (x86) Application")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName "AddCustomerJob"
# PROP Scc_LocalPath "."
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "AddCustomerJob - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /Gm /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /W3 /Gm /GX /ZI /Od /I "C:\Program Files\Intuit\QuickBooks Pro" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /FR /Yu"stdafx.h" /FD /GZ /c
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /debug /machine:I386 /out:"Debug/AddCustomerJob.exe"
# Begin Custom Build - Performing registration
OutDir=.\Debug
TargetPath=.\AddCustomerJob.exe
InputPath=.\AddCustomerJob.exe
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	"$(TargetPath)" /RegServer 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	echo Server registration done! 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "AddCustomerJob - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /GX /O2 /I "C:\Program Files\Intuit\QuickBooks Pro" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /FR /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /W3 /GX /O2 /I "C:\Program Files\Intuit\QuickBooks Pro" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /FR /Yu"stdafx.h" /FD /c
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /machine:I386 /out:"Release/AddCustomerJob.exe"
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /machine:I386 /out:"Release/AddCustomerJob.exe"
# Begin Custom Build - Performing registration
OutDir=.\Release
TargetPath=.\Release\AddCustomerJob.exe
InputPath=.\Release\AddCustomerJob.exe
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	"$(TargetPath)" /RegServer 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	echo Server registration done! 
	
# End Custom Build

!ENDIF 

# Begin Target

# Name "AddCustomerJob - Win32 Debug"
# Name "AddCustomerJob - Win32 Release"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\AddCustomerJob.cpp
# End Source File
# Begin Source File

SOURCE=.\AddCustomerJob.idl
# ADD MTL /tlb ".\AddCustomerJob.tlb" /h "AddCustomerJob.h" /iid "AddCustomerJob_i.c" /Oicf
# End Source File
# Begin Source File

SOURCE=.\AddCustomerJob.rc
# End Source File
# Begin Source File

SOURCE=.\DOMXMLBuilder.cpp
# End Source File
# Begin Source File

SOURCE=.\QBAddCustomerJobDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\QBCompanyFileDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\QBCustomerAdd.cpp
# End Source File
# Begin Source File

SOURCE=.\QBCustomerQuery.cpp
# End Source File
# Begin Source File

SOURCE=.\QBCustomerRet.cpp
# End Source File
# Begin Source File

SOURCE=.\QBCustomerTypeAdd.cpp
# End Source File
# Begin Source File

SOURCE=.\QBCustomerTypeQuery.cpp
# End Source File
# Begin Source File

SOURCE=.\QBCustomerTypeRet.cpp
# End Source File
# Begin Source File

SOURCE=.\qbXMLOpsBase.cpp
# End Source File
# Begin Source File

SOURCE=.\qbXMLRPWrapper.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# Begin Source File

SOURCE=.\Utility.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\DOMXMLBuilder.h
# End Source File
# Begin Source File

SOURCE=.\QBAddCustomerJobDlg.h
# End Source File
# Begin Source File

SOURCE=.\QBCompanyFileDlg.h
# End Source File
# Begin Source File

SOURCE=.\QBCustomerAdd.h
# End Source File
# Begin Source File

SOURCE=.\QBCustomerQuery.h
# End Source File
# Begin Source File

SOURCE=.\QBCustomerRet.h
# End Source File
# Begin Source File

SOURCE=.\QBCustomerTypeAdd.h
# End Source File
# Begin Source File

SOURCE=.\QBCustomerTypeQuery.h
# End Source File
# Begin Source File

SOURCE=.\QBCustomerTypeRet.h
# End Source File
# Begin Source File

SOURCE=.\qbXMLOpsBase.h
# End Source File
# Begin Source File

SOURCE=.\qbXMLRPWrapper.h
# End Source File
# Begin Source File

SOURCE=.\qbXMLTags.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# Begin Source File

SOURCE=.\Utility.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\AddCustomerJob.rgs
# End Source File
# Begin Source File

SOURCE=.\customers.bmp
# End Source File
# Begin Source File

SOURCE=.\qbbanner.bmp
# End Source File
# End Group
# End Target
# End Project
# Section AddCustomerJob : {00000000-0000-0000-0000-800000800000}
# 	1:23:IDD_QBSELECTCUSTOMERDLG:103
# End Section
