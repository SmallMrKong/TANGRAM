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
#include "TangramView_i.h"
#include "dllmain.h"

CCloudAddinDesignerApp theApp;

BOOL CCloudAddinDesignerApp::InitInstance()
{
	WNDCLASS wndClass;
	wndClass.style = CS_DBLCLKS;
	wndClass.lpfnWndProc = DefWindowProc;
	wndClass.cbClsExtra = 0;
	wndClass.cbWndExtra = 0;
	wndClass.hInstance = AfxGetInstanceHandle();
	wndClass.hIcon = 0;
	wndClass.hCursor = ::LoadCursor(NULL, IDC_ARROW);
	wndClass.hbrBackground = 0;
	wndClass.lpszMenuName = NULL;
	wndClass.lpszClassName = _T("Tangram Designer Class");

	if (!AfxRegisterClass(&wndClass))
	{
	}
	CComPtr<ITangram> pCore;
	pCore.CoCreateInstance(CComBSTR(L"tangram.tangram.1"));
	theApp.m_pTangram = pCore.Detach();
	return CWinApp::InitInstance();
}
