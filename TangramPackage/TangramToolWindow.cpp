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
#include "CommonIncludes.h"
#include "TangramToolWindow.h"
#include "TangramVSIApp.h"

const GUID& CAppXmlToolWnd::GetToolWindowGuid() const
{
	return CLSID_guidPersistanceSlot;
}

void CAppXmlToolWnd::PostCreate()
{
	CComVariant srpvt;
	srpvt.vt = VT_I4;
	srpvt.intVal = IDB_IMAGES;
	// We don't want to make the window creation fail only becuase we can not set
	// the icon, so we will not throw if SetProperty fails.
	if (SUCCEEDED(GetIVsWindowFrame()->SetProperty(VSFPROPID_BitmapResource, srpvt)))
	{
		srpvt.intVal = 1;
		GetIVsWindowFrame()->SetProperty(VSFPROPID_BitmapIndex, srpvt);
	}

	if (theApp.m_pTangram&&m_pPane)
	{
		HWND h = m_pPane->m_hWnd;

		//Begin add By Tangram Team: Step 2
		theApp.m_pTangram->CreateWndPage((LONGLONG)::GetParent(h), &m_pPage);
		if (m_pPage)
		{
			//CComBSTR strAppStartFilePath(L"res://StartupPage.dll/start.htm");
			CString strPath = _T("res://") + theApp.m_strModulePath + _T("StartupPage.dll/start.htm");
			//CComBSTR strAppStartFilePath(L"http://tangramcloud.com/static/startup.html");
			m_pPage->put_URL(strPath.AllocSysString());
		}
	}
	//End add By Tangram Team!
	m_pPane->m_pAppXmlToolWnd = this;
}

// Function called by VsWindowPaneFromResource at the end of ClosePane.
// Perform necessary cleanup.
void CAppXmlWndPane::PostClosed()
{
	if (nullptr != m_hBackground)
	{
		::DeleteBrush(m_hBackground);
		m_hBackground = nullptr;
	}

	//if (::IsWindow(m_hWnd)&&theApp.m_pTangram&&m_pAppXmlToolWnd->m_pPage)
	//{
	//	m_pAppXmlToolWnd->m_pPage->get_Frame(CComVariant((long)0),&m_pAppXmlToolWnd->m_pFrame);
	//	//CWindow button(this->GetDlgItem(IDC_CLICKME_BTN));
	//	//m_pAppXmlToolWnd->m_pFrame->ModifyHost((long)button.m_hWnd);
	//	CComVariant var;
	//	theApp.m_pTangram->get_AppKeyValue(CComBSTR(_T("HostWnd")), &var);
	//	HWND _hWnd = (HWND)var.llVal;
	//	theApp.m_pTangram->get_AppKeyValue(CComBSTR(_T("HostChildWnd")), &var);
	//	HWND _hChildWnd = (HWND)var.llVal;
	//	if (::IsWindow(_hWnd))
	//	{
	//		m_pAppXmlToolWnd->m_pFrame->ModifyHost((long)_hChildWnd);
	//		::PostMessage(_hWnd, WM_DESTROY, 0, 0);
	//		//::DestroyWindow(_hWnd);
	//	}
	//}

	CComPtr<IVsShell> spShell = GetVsSiteCache().GetCachedService<IVsShell, SID_SVsShell>();
	if (nullptr != spShell && VSCOOKIE_NIL != m_BroadcastCookie)
	{
		spShell->UnadviseBroadcastMessages(m_BroadcastCookie);
		m_BroadcastCookie = VSCOOKIE_NIL;
	}
}