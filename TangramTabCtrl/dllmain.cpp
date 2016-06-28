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
* http://www.cloudaddin.com
*
********************************************************************************/
// dllmain.cpp : DllMain µÄÊµÏÖ¡£

#include "stdafx.h"
#include "resource.h"
#include "TangramTabCtrl_i.c"
#include "dllmain.h"

CTangramApp theApp;

int CTangramApp::ExitInstance()
{
	CMFCVisualManager * pVisualManager = CMFCVisualManager::GetInstance();
	if (pVisualManager != NULL)
	{
		delete pVisualManager;
	}
	return CWinApp::ExitInstance();
}
