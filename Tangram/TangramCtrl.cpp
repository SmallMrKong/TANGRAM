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
********************************************************************************/

// TangramDoc.cpp : Implementation of CTangramCtrl
#include "stdafx.h"
#include "CloudAddinApp.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "TangramCtrl.h"


// CTangramCtrl

CTangramCtrl::CTangramCtrl()
{
	m_bWindowOnly = true;
	m_hHostWnd = NULL;
	m_hEclipseViewWnd = NULL;
	m_pPage = nullptr;
	m_pCurNode = nullptr;
	m_pEclipseWnd = nullptr;
}

HRESULT CTangramCtrl::FinalConstruct()
{
	return S_OK;
}

void CTangramCtrl::FinalRelease()
{
}

STDMETHODIMP CTangramCtrl::get_HWND(long* pVal)
{
	*pVal = (long)m_hWnd;
	return S_OK;
}

STDMETHODIMP CTangramCtrl::GetCtrlText(long nHandle, BSTR bstrNodeName, BSTR bstrCtrlName, BSTR* bstrVal)
{
	//switch (nHandle)
	//{
	//case 0:
	//	if (m_pFrame)
	//	{
	//		CString strName = OLE2T(bstrNodeName);
	//		CWndFrame* _pFrame = (CWndFrame*)m_pFrame;
	//		auto it = _pFrame->m_mapNode.find(strName);
	//		if (it != _pFrame->m_mapNode.end())
	//		{
	//			CWndNode* pNode = it->second;
	//			if (pNode->m_nViewType == CLRCtrl)
	//			{
	//				CWndPage* pPage = _pFrame->m_pPage;
	//				pPage->GetCtrlValueByName(pNode->m_pDisp, bstrCtrlName, true, bstrVal);
	//			}
	//		}
	//	}
	//	break;
	//case 1:
	//	break;
	//default:
	//	break;
	//}

	return S_OK;
}

STDMETHODIMP CTangramCtrl::InitCtrl(BSTR bstrXml)
{
	if (theApp.m_pJavaProxy)
	{
		theApp.m_pJavaProxy->InitComProxy();
	}
	CString strXml = OLE2T(bstrXml);
	if (strXml != _T(""))
	{
		CTangramXmlParse m_Parse;
		if (!m_Parse.LoadXml(strXml))
		{
			CString strPath = theApp.m_strAppPath + strXml;
			if (m_Parse.LoadFile(strPath) == false)
				return S_OK;
		}
		int nCount = m_Parse.GetCount();
		CTangramXmlParse* pChild = nullptr;
		for (int i = 0; i < nCount; i++)
		{
			pChild = m_Parse.GetChild(i);
			CString strName = pChild->name();
			strName.MakeLower();
			auto it = m_mapTangramInfo.find(strName);
			if (it == m_mapTangramInfo.end())
			{
				m_mapTangramInfo[strName] = pChild->xml();
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CTangramCtrl::put_TangramHandle(BSTR bstrHandleName, LONG newVal)
{
	CString strName = OLE2T(bstrHandleName);
	HWND hWnd = (HWND)newVal;
	if (strName != _T("")&&::IsWindow(hWnd))
	{
		m_hHostWnd = m_hWnd;
		auto it = m_mapTangramHandle.find(strName);
		if (it != m_mapTangramHandle.end())
			m_mapTangramHandle.erase(strName);
		m_mapTangramHandle[strName] = hWnd;
		if (strName.Compare(_T("EclipseView")) == 0)
		{
			m_hEclipseViewWnd = hWnd;
			HWND hTop = ::GetAncestor(m_hEclipseViewWnd, GA_ROOT);

			EclipseCloudPlus::EclipsePlus::CEclipseCloudAddin* pAddin = (EclipseCloudPlus::EclipsePlus::CEclipseCloudAddin*)theApp.m_pHostCore;
			auto it = pAddin->find(hTop);
			if (it == pAddin->end())
			{
				m_pEclipseWnd = new CComObject<CEclipseWnd>;
				m_pEclipseWnd->SubclassWindow(hTop);
			}
			else
			{
				m_pEclipseWnd = it->second;
			}
			auto it2 = m_pEclipseWnd->find(m_hWnd);
			if (it2 == m_pEclipseWnd->end())
				(*m_pEclipseWnd)[m_hWnd] = this;
			HWND hClient = ::GetWindow(hTop, GW_CHILD);
			hClient = ::GetWindow(hClient, GW_CHILD);
			m_pEclipseWnd->m_hClient = hClient;
			m_mapTangramHandle[_T("TopView")] = hClient;
		}
	}

	return S_OK;
}


STDMETHODIMP CTangramCtrl::Extend(BSTR bstrFrameName, BSTR bstrKey, BSTR bstrXml, IWndNode** ppNode)
{
	if (::IsWindow(m_hEclipseViewWnd))
	{
		CString strFrameName = OLE2T(bstrFrameName);
		if (strFrameName == _T(""))
			return S_OK;
		if (strFrameName.CompareNoCase(_T("TopView")) == 0&& m_pEclipseWnd)
		{
			return m_pEclipseWnd->SWTExtend(bstrKey, bstrXml, ppNode);
		}
		auto it = m_mapTangramHandle.find(strFrameName);
		if (it == m_mapTangramHandle.end())
			return S_OK;
		HWND hWnd = it->second;
		CString strKey = OLE2T(bstrKey);
		if (strKey == _T(""))
			return S_OK;
		CString strXml = OLE2T(bstrXml);
		if (strXml == _T(""))
			return S_OK;
		if (m_pPage == nullptr)
		{
			theApp.m_pTangram->CreateWndPage((long)::GetParent(m_hEclipseViewWnd), &m_pPage);
			if (m_pPage == nullptr)
				return S_FALSE;
		}

		auto it2 = m_mapTangramFrame.find(strFrameName);
		if (it2==m_mapTangramFrame.end())
		{
			IWndFrame* pFrame = nullptr;
			m_pPage->CreateFrame(CComVariant(0), CComVariant((long)hWnd), bstrFrameName, &pFrame);
			if (pFrame == nullptr)
			{
				delete m_pPage;
				m_pPage = nullptr;
				return S_OK;
			}
			m_mapTangramFrame[strFrameName] = pFrame;
			return pFrame->Extend(bstrKey, bstrXml, ppNode);
		}
		return it2->second->Extend(bstrKey, bstrXml, ppNode);
	}
	return S_OK;
}

void CTangramCtrl::OnFinalMessage(HWND hWnd)
{
	if (m_pEclipseWnd)
	{
		auto it = m_pEclipseWnd->find(m_hHostWnd);
		if (it != m_pEclipseWnd->end())
			m_pEclipseWnd->erase(it);
	}
	__super::OnFinalMessage(hWnd);
}
