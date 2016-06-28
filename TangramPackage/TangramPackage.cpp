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
* mailto:sunhuizlz@yeah.net
* http://www.CloudAddin.com
*
*
********************************************************************************/

#include "stdafx.h"
#include "TangramVSIApp.h"
#include "CommonIncludes.h"

#pragma comment(lib,"Version.lib")
#pragma comment(lib,"Wininet.lib")
#pragma comment(lib,"Winmm.lib")

CAppDocPackage::CAppDocPackage() :
	m_dwEditorCookie(0),
	m_ToolWindow(GetVsSiteCache())
{
	theApp.m_pTangramPackage = this;
}

CAppDocPackage::~CAppDocPackage()
{
}

// This method will be called after IVsPackage::SetSite is called with a valid site
void CAppDocPackage::PostSited(IVsPackageEnums::SetSiteResult /*result*/)
{
	if (m_dwEditorCookie == 0)
	{
		CComObject<CAppXmlDocEditorFactory> *pFactory = new CComObject<CAppXmlDocEditorFactory>;
		if (NULL == pFactory)
		{
			ERRHR(E_OUTOFMEMORY);
		}
		CComPtr<IVsEditorFactory> spIVsEditorFactory = static_cast<IVsEditorFactory*>(pFactory);
		// Register the editor factory
		CComPtr<IVsRegisterEditors> spIVsRegisterEditors;
		CHKHR(GetVsSiteCache().QueryService(SID_SVsRegisterEditors, &spIVsRegisterEditors));
		CHKHR(spIVsRegisterEditors->RegisterEditor(CLSID_TangramPackageEditorFactory, spIVsEditorFactory, &m_dwEditorCookie));
	}
	if (theApp.m_pDTE == NULL)
	{
		CreateTool(CLSID_guidPersistanceSlot);
		CHKHR(GetVsSiteCache().QueryService(SID_SDTE, &theApp.m_pDTE));
		if (theApp.m_pDTE)
			theApp.InitInstance();
	}
}

void CAppDocPackage::PreClosing()
{
	if (m_dwEditorCookie != 0)
	{
		if (theApp.m_pTangram)
		{
			//HWND _hWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), L"", WS_CHILD, 0, 0, 0, 0, theApp.m_pHostCore->m_hHostWnd, NULL, AfxGetInstanceHandle(), NULL);
			//HWND _hChildWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), L"", WS_CHILD, 0, 0, 0, 0, (HWND)_hWnd, NULL, AfxGetInstanceHandle(), NULL);
			//if (::IsWindow(m_hWnd))
			//{
			//	m_pTaskPaneFrame->ModifyHost((long)_hChildWnd);
			//	::DestroyWindow(m_hWnd);
			//}
			theApp.m_pTangram.Detach();
		}
		CComPtr<IVsRegisterEditors> spIVsRegisterEditors;
		CHKHR(GetVsSiteCache().QueryService(SID_SVsRegisterEditors, &spIVsRegisterEditors));
		CHKHR(spIVsRegisterEditors->UnregisterEditor(m_dwEditorCookie));
	}
}

void CAppDocPackage::CreateTangramToolWnd()
{
	m_ToolWindow.CreateAndShow();
}
