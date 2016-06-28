/********************************************************************************
*					Tangram Library - version 8.0								*
*********************************************************************************
* Copyright (C) 2002-2016 by Tangram Team.   All Rights Reserved.				*
*
* THIS SOURCE FILE IS THE PROPERTY OF TANGRAM TEAM AND IS NOT TO
* BE RE-DISTRIBUTED BY ANY MEANS WHATSOEVER WITHOUT THE EXPRESSED
* WRITTEN CONSENT OF TANGRAM TEAM.
*
* THIS SOURCE CODE CAN ONLY BE USED UNDER THE TERMS AND CONDITIONS
* OUTLINED IN THE GPL LICENSE AGREEMENT.TANGRAM TEAM
* GRANTS TO YOU (ONE SOFTWARE DEVELOPER) THE LIMITED RIGHT TO USE
* THIS SOFTWARE ON A SINGLE COMPUTER.
*
* CONTACT INFORMATION:
* mailto:sunhuizlz@yeah.net, mailto:sunhui@cloudaddin.com
* http://www.CloudAddin.com
*
*
********************************************************************************/
// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently


#pragma once

// Windows Platform headers and control defines
#define STRICT
#define _WIN32_WINNT 0x0501  // Visual Studio requires Windows 2000 or better
//#define NOMINMAX // Windows Platform min and max macros cause problems for the Standard C++ Library
#define WIN32_LEAN_AND_MEAN	// Exclude rarely-used stuff from the Windows Platform headers

// ATL headers and control defines
#define _ATL_APARTMENT_THREADED
#define _ATL_REGISTER_PER_USER
#include <windows.h>

// Windows Platform headers
#include <windowsx.h> // REVIEW - what is this for?
#include <richedit.h>
#include <TOM.h>

#include <atlbase.h>
#include <atlcom.h>
#include <atlwin.h>
#include <atlstr.h>
#include <atlfile.h>
#include <atlsafe.h>
#include <atlstr.h>

// Visual Studio Platform headers
#include <dte.h> // for extensibility
#include <objext.h> // for ILocalRegistry
#include <vshelp.h> // for Help
#include <uilocale.h> // for IUIHostLocale2
#include <IVsQueryEditQuerySave2.h> // for IVsQueryEditQuerySave2
#include <vbapkg.h> // for IVsMacroRecorder
#include <fpstfmt.h> // for IPersistFileFormat
#include <VSRegKeyNames.h>
#include <stdidcmd.h>
#include <stdiduie.h> // For status bar consts.
#include <textfind.h>
#include <textmgr.h>

#include "tangram.h"
#include "Shlobj.h"

// VSL headers
#define VSLASSERT _ASSERTE
#define VSLASSERTEX(exp, szMsg) _ASSERT_BASE(exp, szMsg)
#define VSLTRACE ATLTRACE
#define FORCE_EXPLICIT_DTE_NAMESPACE
#define FORCE_EXPLICIT_PRJ_NAMESPACE

#include <VSLPackage.h>
#include <VSLCommandTarget.h>
//#include <VSLWindows.h>
#include <VSLControls.h>
#include <VSLFile.h>
#include <VSLContainers.h>
#include <VSLComparison.h>
#include <VSLAutomation.h>
#include <VSLFindAndReplace.h>
#include <VSLShortNameDefines.h>

using namespace VSL;
#include "TangramXmlParse.h"
#include <map>

#define	UPDATE_STATUS						-1
#define UPDATE_STATUS_SUCCEED				(UPDATE_STATUS + 1)
#define	UPDATE_STATUS_USER_CANCELED			(UPDATE_STATUS + 2)
#define UPDATE_STATUS_CLOSE_HOST_FAILUR		(UPDATE_STATUS + 3)
#define UPDATE_STATUS_PARSE_XMLCFG_FAILUR	(UPDATE_STATUS + 4)
#define UPDATE_STATUS_START_HOST_FAILUR		(UPDATE_STATUS + 5)

using namespace ATL;

