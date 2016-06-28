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
* http://www.cloudaddin.com
*
********************************************************************************/
// TangramConnector.cpp : CTangramConnector 的实现

#include "stdafx.h"
#include "TangramConnector.h"
#include "TangramTabCtrlWnd.h"
#include "TangramTabCtrl_i.c"
#include "Tangram.c"

// CTangramConnector
STDMETHODIMP CTangramConnector::Create(long hParent, IWndNode* pNode, long* hWnd)
{
	BSTR bstrTag = L"";
	BSTR bstrStyle = L"";

	USES_CONVERSION;

	pNode->get_Attribute(L"id", &bstrTag);
	pNode->get_Attribute(L"style", &bstrStyle);

	HWND hPWnd = (HWND)hParent;
	CString m_strTag = OLE2T(bstrTag);
	::SysFreeString(bstrTag);
	HRESULT hr = S_OK;

	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	m_strTag.Trim();
	if (::IsWindow(hPWnd))
	{
		//RECT rt;
		//::GetClientRect(hPWnd, &rt);
		//if (m_strTag.CompareNoCase(_T("TestView")) == 0)
		//{
		//	CRuntimeClass* pRC = RUNTIME_CLASS(CTangramFormView);
		//	CTangramFormView* pTangramFormView = (CTangramFormView*)pRC->CreateObject();
		//	if (pTangramFormView)
		//	{
		//		CCreateContext context;
		//		pTangramFormView->Create(NULL, _T(""), WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, CRect(0, 0, 0, 0), CWnd::FromHandle(hPWnd), NULL, &context);
		//		*hWnd = (LONGLONG)pTangramFormView->m_hWnd;
		//	}
		//	
		//	return S_OK;
		//}
		if (m_strTag.CompareNoCase(_T("TangramTabCtrl")) == 0)
		{
			BSTR bstrLocation = L"";
			pNode->get_Attribute(L"Location", &bstrLocation);
			CMFCTabCtrl::Location loc = (CMFCTabCtrl::Location)_ttoi(OLE2T(bstrStyle));
			CTangramTabCtrlWnd* pTangramTabCtrlWnd = new CTangramTabCtrlWnd();
			pTangramTabCtrlWnd->m_pWndNode = pNode;
			if (pTangramTabCtrlWnd)
			{
				CRect rectDummy;
				rectDummy.SetRectEmpty();
				CMFCTabCtrl::Style nStyle = (CMFCTabCtrl::Style)_ttoi(OLE2T(bstrStyle));
				//::SysFreeString(bstrStyle);
				//::SysFreeString(bstrLocation);
				// 创建选项卡窗口: 
				pTangramTabCtrlWnd->SetLocation(loc);
				if (!pTangramTabCtrlWnd->Create(nStyle, rectDummy, CWnd::FromHandle(hPWnd), 1))
				{
					return -1;      // 未能创建
				}
				*hWnd = (long)pTangramTabCtrlWnd->m_hWnd;
			}
			
			return S_OK;
		}

		return S_OK;
	}
	return S_OK;
}


STDMETHODIMP CTangramConnector::get_Names(BSTR * pVal)
{
	*pVal = ::SysAllocString(L"TangramTabCtrl");

	//if (strKey != _T(""))
	//{
	//	::SendMessage(m_hWnd, WM_TANGRAMMSG, (WPARAM)strKey.GetBuffer(), 0);
	//	strKey.Trim();
	//	strKey.MakeLower();
	//	map<CString, VARIANT>::iterator it = m_mapVal.find(strKey);
	//	if (it != m_mapVal.end())
	//		*pVal = it->second;
	//	strKey.ReleaseBuffer();
	//}
	return S_OK;
}

STDMETHODIMP CTangramConnector::get_Tags(BSTR bstrName, BSTR * pVal)
{
	*pVal = ::SysAllocString(L"0,1,2,3,4,5,6,7,");
	return S_OK;
}