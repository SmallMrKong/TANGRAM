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


#include "../stdafx.h"
#include "../CloudAddinApp.h"
#include "../WndNode.h"
#include "../WndFrame.h"
#include "../TangramCtrl.h"
#include "EclipseAddin.h"

namespace EclipseCloudPlus
{
	namespace EclipsePlus
	{
		CEclipseCloudAddin::CEclipseCloudAddin()
		{
			m_bClose = false;
			m_nIndex = 0;
			m_strURL = _T("");
		}

		CEclipseCloudAddin::~CEclipseCloudAddin()
		{
		}

		void CEclipseCloudAddin::InitTangramforJava()
		{

		}

		void CEclipseCloudAddin::InitEclipsePlus()
		{
			if (m_hHostWnd == NULL)
			{
				if (theApp.m_hCBTHook == NULL)
					theApp.m_hCBTHook = SetWindowsHookEx(WH_CBT, CTangramApp::CBTProc, NULL, ::GetCurrentThreadId());
				theApp.m_bEnableProcessFormTabKey = true;
				m_hHostWnd = ::CreateWindowEx(WS_EX_PALETTEWINDOW, _T("Tangram Window Class"), _T("Tangram Designer"), WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, 0, 0, 0, 0, NULL, NULL, theApp.m_hInstance, NULL);
				m_hChildHostWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), _T(""), WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, m_hHostWnd, NULL, theApp.m_hInstance, NULL);
			}
		}

		CEclipseWnd::CEclipseWnd(void)
		{
			m_nIndex = 0;
			m_bCreated = false;
			m_strURL = _T("");
			m_pCurNode = nullptr;
			m_pPage = nullptr;
			m_pFrame = nullptr;
		}

		CEclipseWnd::~CEclipseWnd(void) 
		{
		}

		void CEclipseWnd::OnFinalMessage(HWND hWnd)
		{
			CWindowImpl::OnFinalMessage(hWnd);
			delete this;
		}

		STDMETHODIMP CEclipseWnd::get_Handle(long* pVal)
		{
			*pVal = (long)m_hClient;
			return S_OK;
		}

		STDMETHODIMP CEclipseWnd::SWTExtend(BSTR bstrKey, BSTR bstrXml, IWndNode** ppNode)
		{
			CString strKey = OLE2T(bstrKey);
			strKey.Trim();

			if (m_hClient == NULL)
				return S_FALSE;
			if (m_pPage == nullptr)
			{
				theApp.m_pTangram->CreateWndPage((long)m_hWnd, &m_pPage);
				if (m_pPage == nullptr)
					return S_FALSE;
				if (m_pFrame == nullptr)
				{
					m_pPage->CreateFrame(CComVariant(0), CComVariant((long)m_hClient), CComBSTR(L"SWT"), &m_pFrame);
					if (m_pFrame == nullptr)
					{
						delete m_pPage;
						m_pPage = nullptr;
					}
				}
			}

			HRESULT hr = m_pFrame->Extend(bstrKey, bstrXml, ppNode);
			if (*ppNode)
			{
				m_pCurNode = *ppNode;
			}
			return hr;
		}

		STDMETHODIMP CEclipseWnd::GetCtrlText(BSTR bstrNodeName, BSTR bstrCtrlName, BSTR* bstrVal)
		{
			if (m_pFrame)
			{
				CString strName = OLE2T(bstrNodeName);
				CWndFrame* _pFrame = (CWndFrame*)m_pFrame;
				auto it = _pFrame->m_mapNode.find(strName);
				if (it != _pFrame->m_mapNode.end())
				{
					CWndNode* pNode = it->second;
					if (pNode->m_nViewType == CLRCtrl)
					{
						CWndPage* pPage = _pFrame->m_pPage;
						//pPage->GetCtrlByName
					}
				}
			}
			return S_OK;
		}

		LRESULT CEclipseWnd::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			CEclipseCloudAddin* pAddin = ((CEclipseCloudAddin*)theApp.m_pHostCore);
			auto it = pAddin->find(m_hWnd);
			if (it != pAddin->end())
				pAddin->erase(it);
			if (pAddin->size() == 0)
			{
				pAddin->m_bClose = true;
			}
			LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
			return lRes;
		}

		LRESULT CEclipseWnd::OnUCMAMsg(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			UCMAMSGInfo* pUCMAMSGInfo = (UCMAMSGInfo*)wParam;
			IWndNode* pNode = nullptr;
			SWTExtend(pUCMAMSGInfo->m_strMsg.AllocSysString(), pUCMAMSGInfo->m_strMsg.AllocSysString(), &pNode);
			delete pUCMAMSGInfo;
			LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
			return lRes;
		}

		CEclipseSWTWnd::CEclipseSWTWnd(void)
		{
			m_pHostCtrl = NULL;
		}

		CEclipseSWTWnd::~CEclipseSWTWnd(void)
		{
		}

		void CEclipseSWTWnd::OnFinalMessage(HWND hWnd)
		{
			CWindowImpl::OnFinalMessage(hWnd);
			delete this;
		}

		LRESULT CEclipseSWTWnd::OnSWTComponentNotify(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			if (m_pHostCtrl)
			{
				LPCTSTR lpstrExtenderID = (LPCTSTR)wParam;
				CString strXML = (LPCTSTR)lParam;
				CComPtr<IDispatch> pDisp;
				m_pHostCtrl->m_pPage->get_Extender(CComBSTR(lpstrExtenderID),&pDisp);
				if (pDisp)
				{
					CComQIPtr<IAppExtender> pAppExtender(pDisp);
					if (pAppExtender)
					{
						pAppExtender->ProcessNotify(strXML.AllocSysString());
					}
				}
			}
			return 0;
		}
	}
}
