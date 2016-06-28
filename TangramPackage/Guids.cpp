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

// Guids.c
//

#include "initguid.h"
#include "Guids.h"
#include <VSShellInterfaces.h>

#define _MIDL_USE_GUIDDEF_ // Necessary to build without compilation errors
#include "TangramPackage.c"

// There is no library that defines the GUIDs for the interfaces defined inside
// TOM.h, so we have to define here the GUIDs that are used inside our code.
// You can find these GUIDs looking at the definition of the interfaces in TOM.h.
DEFINE_GUID(IID_ITextDocument,
    0x8CC497C0, 0xA1DF, 0x11ce, 0x80, 0x98, 0x00, 0xAA, 0x00, 0x47, 0xBE, 0x5D);

