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
// TangramConnector.cpp : CTangramConnector µÄÊµÏÖ

#include "stdafx.h"
#include "TangramConnector.h"
#include "TangramNodeView.h"

// CTangramConnector
STDMETHODIMP CTangramConnector::Create(long hParent, IWndNode* pNode, long* hWnd)
{
	BSTR bstrTag = L"";
	BSTR bstrStyle = L"";
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
		RECT rt;
		::GetClientRect(hPWnd, &rt);
		if (m_strTag.CompareNoCase(_T("TangramNodeView")) == 0)
		{
			CRuntimeClass* pRC = RUNTIME_CLASS(CTangramNodeView);
			CTangramNodeView* pTangramFormView = (CTangramNodeView*)pRC->CreateObject();
			if (pTangramFormView)
			{
				pTangramFormView->m_pNode = pNode;
				CCreateContext context;
				pTangramFormView->Create(NULL, _T(""), WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, CRect(0, 0, 0, 0), CWnd::FromHandle(hPWnd), NULL, &context);
				*hWnd = (long)pTangramFormView->m_hWnd;
				return S_OK;
			}
		}
	}
	return S_OK;
}


STDMETHODIMP CTangramConnector::get_Names(BSTR * pVal)
{
	CString strVal = _T("TangramNodeView,");
	*pVal = strVal.AllocSysString();
	strVal.MakeLower();
	int nPos = strVal.Find(_T(","));
	int nIndex = 0;
	while (nPos != -1)
	{
		CString strKey = strVal.Left(nPos);
		switch (nIndex)
		{
		case 0:
			break;
		case 1:
			break;
		case 2:
			break;
		case 3:
			break;
		case 4:
			break;
		case 5:
			break;
		case 6:
			break;
		case 7:
			break;
		case 8:
			break;
		case 9:
			break;
		}
		nIndex++;
		strVal = strVal.Mid(nPos + 1);
		nPos = strVal.Find(_T(","));
	}

	return S_OK;
}

STDMETHODIMP CTangramConnector::get_Tags(BSTR bstrName, BSTR * pVal)
{
	CString strKey = OLE2T(bstrName);
	strKey.MakeLower().Trim();
	auto it = m_mapIndex.find(strKey);
	if (it != m_mapIndex.end())
	{
		*pVal = it->second.AllocSysString();
	}

	return S_OK;
}
