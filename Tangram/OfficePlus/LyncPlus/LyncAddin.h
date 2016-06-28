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
* OUTLINED IN THE TANGRAM LICENSE AGREEMENT.TANGRAM TEAM
* GRANTS TO YOU (ONE SOFTWARE DEVELOPER) THE LIMITED RIGHT TO USE
* THIS SOFTWARE ON A SINGLE COMPUTER.
*
* CONTACT INFORMATION:
* mailto:sunhuizlz@yeah.net
* http://www.CloudAddin.com
*
********************************************************************************/


#pragma once
#include "CommonFunction.h"
#include "def.h"
#include "uccapievent.h"
#include "LyncEvent.h"

#include "..\Tangram\CloudAddinCore.h"

using namespace UCCollaborationLib;
using namespace OfficeCloudPlus::LyncPlus::LyncClientEvent;

namespace OfficeCloudPlus
{
	namespace LyncPlus
	{
		class CLyncCloudAddin : 
			public CTangram,
			public CTangramLyncClientEvents
		{
		public:
			CLyncCloudAddin();
			virtual ~CLyncCloudAddin();
		};
	}
}


