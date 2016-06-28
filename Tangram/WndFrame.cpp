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

// WndFrame.cpp : implementation file of CWndFrame
//

#include "stdafx.h"
#include "CloudAddinApp.h"
#include "CloudAddinCore.h"
#include "NodeWnd.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "TangramHtmlTreeWnd.h"
#include "OfficePlus\WordPlus\WordAddin.h"

using namespace OfficeCloudPlus::WordPlus;

// CWndFrame
CWndFrame::CWndFrame()
{
	m_x					 = 0;
	m_y					 = 0;
	m_left				 = 0;
	m_top				 = 0;
	m_right				 = 0;
	m_bottom			 = 0;
	m_left2				 = 0;
	m_top2				 = 0;
	m_right2			 = 0;
	m_bottom2			 = 0;
	m_hLWnd				 = 0;
	m_hTWnd				 = 0;
	m_hRWnd				 = 0;
	m_hBWnd				 = 0;
	m_hOuterLWnd		 = 0;
	m_hOuterTWnd		 = 0;
	m_hOuterRWnd		 = 0;
	m_hOuterBWnd		 = 0;
	m_strCurrentKey		 = _T("");
	m_strFrameName		 = _T("");
	m_bDetached			 = false;
	m_bDesignerState	 = false;
	m_bOuterInited		 = false;
	m_bNeedOuterChanged	 = true;
	m_pBKWnd			 = nullptr;
	m_pPage				 = nullptr;
	m_pHostNode			 = nullptr;
	m_pWorkNode			 = nullptr;
	m_pRootNodes		 = nullptr;
	m_pBindingNode		 = nullptr;
#ifdef _DEBUG
	theApp.m_nTangramFrame++;
#endif
}

CWndFrame::~CWndFrame()
{
#ifdef _DEBUG
	theApp.m_nTangramFrame--;
#endif	
	if (theApp.m_pWndFrame == this)
		theApp.m_pWndFrame = nullptr;
	if (m_pRootNodes)
		CCommonFunction::ClearObject<CWndNodeCollection>(m_pRootNodes);
	if (m_mapVal.size())
	{
		for (auto it : m_mapVal)
		{
			::VariantClear(&it.second);
		}
		m_mapVal.clear();
	}
	if (m_pPage)
	{
		auto it = m_pPage->m_mapFrame.find(m_hHostWnd);
		if (it != m_pPage->m_mapFrame.end())
		{
			m_pPage->m_mapFrame.erase(it);
			if (m_pPage->m_mapFrame.size()==0)
				delete m_pPage;
		}
	}
	m_hWnd = NULL; 
}

void CWndFrame::UpdateWndNode()
{
	for (auto it : m_mapNode)
	{
		CWndNode* pWindowNode = (CWndNode*)it.second;
		if (pWindowNode)
		{
			CComQIPtr<IWndContainer> pContainer(pWindowNode->m_pDisp);
			if (pContainer)
			{
				if (pWindowNode->m_nActivePage >0)
				{
					CString strVal = _T("");
					strVal.Format(_T("%d"), pWindowNode->m_nActivePage);
					pWindowNode->m_pHostParse->put_attr(_T("activepage"), strVal);
				}
				pContainer->Save();
			}

			for (auto it2 : pWindowNode->m_vChildNodes)
			{
				theApp.m_pTangram->UpdateWndNode(it2);
			}

			if (pWindowNode->m_pOfficeObj)
			{
				theApp.m_pHostCore->UpdateOfficeObj(pWindowNode->m_pOfficeObj, pWindowNode->m_pTangramParse->GetChild(_T("window"))->xml(), pWindowNode->m_pTangramParse->name());
			}
		}
	}
}

void CWndFrame::UpdateDesignerTreeInfo()
{
	if (m_bDesignerState&&theApp.m_pHostCore->m_hChildHostWnd)
	{
		theApp.m_pDesigningFrame = this;
		if (theApp.m_pDocDOMTree)
		{
			theApp.m_pDocDOMTree->DeleteItem(theApp.m_pDocDOMTree->m_hFirstRoot);
			if (theApp.m_pDocDOMTree->m_pHostXmlParse)
			{
				delete theApp.m_pDocDOMTree->m_pHostXmlParse;
				theApp.m_pDocDOMTree->m_pHostXmlParse = nullptr;
			}
			CWndNode* pNode = theApp.m_pDesigningFrame->m_pWorkNode;
			if (pNode == nullptr)
			{
				return;
			}
			CString strName = pNode->m_strName;
			CString _strName = strName;
			_strName += _T("-indesigning");
			_strName.MakeLower();
			CTangramXmlParse* pParse = nullptr;
			auto it = m_mapNode.find(_strName);
			if (it != m_mapNode.end())
				pParse = it->second->m_pTangramParse;
			else
				pParse = pNode->m_pTangramParse;
			theApp.InitDesignerTreeCtrl(pParse->xml());
		}
	}
}

CWndNode* CWndFrame::OpenXtmlDocument(CString strKey, CString strFile)
{	
	m_pWorkNode = new CComObject<CWndNode>;
	m_pWorkNode->m_pTangramParse = theApp.m_pCurWorkNodeParse;
	CTangramXmlParse* pParse = theApp.m_pCurWorkNodeParse->GetChild(_T("window"));
	m_pWorkNode->m_pHostParse = pParse->GetChild(_T("node"));
	theApp.m_pCurWorkNodeParse = nullptr;
	m_pWorkNode->m_pRootObj = m_pWorkNode;
	m_pWorkNode->m_pFrame = this;

	CreateWndPage();
	m_mapNode[strKey] = m_pWorkNode;

	if (m_pPage)
		m_pPage->Fire_ExtendXmlComplete(strFile.AllocSysString(), (long)m_hHostWnd, m_pWorkNode);
	m_pWorkNode->m_strKey = strKey;
	m_pWorkNode->Fire_ExtendComplete();

	return m_pWorkNode;
}

BOOL CWndFrame::CreateWndPage()
{
	if(::IsWindow(m_hWnd)==false)
		SubclassWindow(m_hHostWnd);
	CWnd* pParentWnd = CWnd::FromHandle(::GetParent(m_hWnd));
	CComPtr<IXMLDOMElement> pRootElement;

	CString strCaption = m_pWorkNode->m_pTangramParse->attr(TGM_CAPTION, _T(""));
	if(m_left==0)
		m_left = m_pWorkNode->m_pTangramParse->attrInt(_T("DeltLeft"),0);
	if(m_top==0)
		m_top = m_pWorkNode->m_pTangramParse->attrInt(_T("delttop"),0);
	if(m_right==0)
		m_right = m_pWorkNode->m_pTangramParse->attrInt(_T("deltright"),0);
	if(m_bottom==0)
		m_bottom = m_pWorkNode->m_pTangramParse->attrInt(_T("deltbottom"),0);
	if(m_left2==0)
		m_left2 = m_pWorkNode->m_pTangramParse->attrInt(_T("deltleft2"),0);
	if(m_top2==0)
		m_top2 = m_pWorkNode->m_pTangramParse->attrInt(_T("delttop2"),0);
	if(m_right2==0)
		m_right2 = m_pWorkNode->m_pTangramParse->attrInt(_T("deltright2"),0);
	if(m_bottom2==0)
		m_bottom2 = m_pWorkNode->m_pTangramParse->attrInt(_T("deltbottom2"),0);

	CTangramXmlParse* pWndParse = m_pWorkNode->m_pTangramParse->GetChild(_T("window"));

	m_pWorkNode->m_strName.Trim();
	m_pWorkNode->m_strName.MakeLower();
	theApp.InitWndNode(m_pWorkNode);
	HWND hWnd = NULL;
	if (m_pWorkNode->m_pObjClsInfo)
	{
		RECT rc;
		m_pWorkNode->m_pRootObj = m_pWorkNode;
		CCreateContext	m_Context;
		m_Context.m_pNewViewClass = m_pWorkNode->m_pObjClsInfo;
		CWnd* pWnd = (CWnd*)m_pWorkNode->m_pObjClsInfo->CreateObject();
		GetWindowRect(&rc);
		if(pParentWnd)
			pParentWnd->ScreenToClient(&rc);

		pWnd->Create(NULL,_T(""),WS_CHILD|WS_VISIBLE,rc,pParentWnd,0,&m_Context);
		hWnd = pWnd->m_hWnd;
		pWnd->ModifyStyle(0,WS_CLIPSIBLINGS);
		if (m_pWorkNode->m_nViewType != Splitter)
			pWnd->ModifyStyleEx(0,WS_EX_CLIENTEDGE);
		else
			pWnd->ModifyStyleEx(WS_EX_CLIENTEDGE,0);
		::SetWindowPos(pWnd->m_hWnd, HWND_BOTTOM,rc.left,rc.top,rc.right-rc.left,rc.bottom-rc.top,SWP_DRAWFRAME|SWP_SHOWWINDOW|SWP_NOACTIVATE);//|SWP_NOREDRAWSWP_NOZORDER);
	}

	pWndParse = m_pWorkNode->m_pTangramParse->GetChild(_T("docplugin"));
	if (pWndParse)
	{
		CString strPlugID = _T("");
		HRESULT hr = S_OK;
		BOOL bHavePlugin = false;
		int nCount = pWndParse->GetCount();
		for (int i = 0; i < nCount; i++)
		{
			CTangramXmlParse* pChild = pWndParse->GetChild(i);
			CComBSTR bstrXml(pChild->xml());
			strPlugID = pChild->text();
			strPlugID.Trim();
			strPlugID.MakeLower();
			if(strPlugID!=_T(""))
			{
				int nPos = strPlugID.Find(_T(","));
				CComBSTR bstrPlugIn(strPlugID);
				CComPtr<IDispatch> pDisp;
				//for COM Component:
				if(nPos==-1)
				{
					hr = pDisp.CoCreateInstance(strPlugID.AllocSysString());
					if(hr==S_OK)
					{
						m_pWorkNode->m_PlugInDispDictionary[strPlugID] = pDisp.p;
						pDisp.p->AddRef();
					}
				
					m_pWorkNode->Fire_NodeAddInCreated(pDisp.p, bstrPlugIn, bstrXml);
				}
				else //for .NET Component
				{
					hr = theApp.m_pTangram->CreateCLRObj(bstrPlugIn,&pDisp);
					if(hr == S_OK)
					{
						m_pWorkNode->m_PlugInDispDictionary[strPlugID] = pDisp.p;
					
						bstrPlugIn = strPlugID.AllocSysString();
						m_pWorkNode->Fire_NodeAddInCreated(pDisp, bstrPlugIn, bstrXml);
					}
				}
				if (m_pPage&&pDisp)
					m_pPage->Fire_AddInCreated(m_pWorkNode, pDisp, bstrPlugIn, bstrXml);
				::SysFreeString(bstrPlugIn);
				bHavePlugin = true;
			}
			::SysFreeString(bstrXml);
		}

		if (bHavePlugin)
			m_pWorkNode->Fire_NodeAddInsCreated();
	}
	m_pWorkNode->m_bCreated = true;
	return true;
}

STDMETHODIMP CWndFrame::get_RootNodes(IWndNodeCollection** pNodeColletion)
{	
	if (m_pRootNodes == nullptr)
	{
		CComObject<CWndNodeCollection>::CreateInstance(&m_pRootNodes);
		m_pRootNodes->AddRef();
	}

	m_pRootNodes->m_pNodes->clear();
	for (auto it : m_mapNode)
	{
		m_pRootNodes->m_pNodes->push_back(it.second);
	}

	return m_pRootNodes->QueryInterface(IID_IWndNodeCollection,(void**)pNodeColletion);
}

STDMETHODIMP CWndFrame::get_FrameData(BSTR bstrKey, VARIANT* pVal)
{
	CString strKey = OLE2T(bstrKey);

	if (strKey != _T(""))
	{
		::SendMessage(m_hWnd, WM_TANGRAMMSG, (WPARAM)strKey.GetBuffer(), 0);
		strKey.Trim();
		strKey.MakeLower();
		auto it = m_mapVal.find(strKey);
		if (it != m_mapVal.end())
			*pVal = it->second;
		strKey.ReleaseBuffer();
	}
	return S_OK;
}

STDMETHODIMP CWndFrame::put_FrameData(BSTR bstrKey, VARIANT newVal)
{
	CString strKey = OLE2T(bstrKey);

	if (strKey == _T("")/*||strVal==_T("")*/)
		return S_OK;
	strKey.Trim();
	strKey.MakeLower();
	m_mapVal[strKey] = newVal;
	return S_OK;
}

STDMETHODIMP CWndFrame::Detach(void)
{
	if(::IsWindow(m_hWnd))
	{
		m_bDetached = true;
		::ShowWindow(m_pWorkNode->m_pHostWnd->m_hWnd,SW_HIDE);
		RECT rect;
		m_pWorkNode->m_pHostWnd->GetWindowRect(&rect);
		m_hHostWnd = UnsubclassWindow(true);
		if(::IsWindow(m_hHostWnd))
		{
			HWND m_hParentWnd = ::GetParent(m_hHostWnd);
			CWnd* pParent = CWnd::FromHandle(m_hParentWnd);
			pParent->ScreenToClient(&rect);
			::SetWindowPos(m_hHostWnd,NULL,rect.left,rect.top,rect.right-rect.left,rect.bottom-rect.top,SWP_FRAMECHANGED|SWP_SHOWWINDOW);
		}
	}
	return S_OK;
}

STDMETHODIMP CWndFrame::Attach(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if(!::IsWindow(m_hWnd))
	{
		m_bDetached = false;
		::ShowWindow(m_pWorkNode->m_pHostWnd->m_hWnd,SW_SHOW);
		SubclassWindow(m_hHostWnd);
	}
	theApp.HostPosChanged(m_pBindingNode);
	return S_OK;
}

STDMETHODIMP CWndFrame::ModifyHost(long hHostWnd)
{
	HWND _hHostWnd = (HWND)hHostWnd;
	if (!::IsWindow(_hHostWnd)||m_hWnd==_hHostWnd)
	{
		return S_OK;
	}
	if (m_pWorkNode == nullptr)
		return S_FALSE;
	HWND hParent = (HWND)::GetParent(_hHostWnd);
	CWindow m_Parent;
	RECT rc;
	m_pWorkNode->m_pHostWnd->GetWindowRect(&rc);
	if(::IsWindow(m_hWnd))
	{
		HWND hTangramWnd = m_pPage->m_hWnd;
		auto it = theApp.m_mapWindowPage.find(hTangramWnd);
		if (it != theApp.m_mapWindowPage.end())
			theApp.m_mapWindowPage.erase(it);
		m_pPage->m_hWnd = hTangramWnd;
		theApp.m_mapWindowPage[hTangramWnd] = m_pPage;

		map<HWND, CWndFrame*>::iterator it2 = m_pPage->m_mapFrame.find(m_hWnd);
		if (it2 != m_pPage->m_mapFrame.end())
		{
			m_pPage->m_mapFrame.erase(it2);
			m_pPage->m_mapFrame[_hHostWnd] = this;
			m_pPage->m_mapWnd[m_strFrameName] = _hHostWnd;
			DWORD dwID = ::GetWindowThreadProcessId(_hHostWnd, NULL);
			TRACE(_T("ExtendEx ThreadInfo:%x\n"), dwID);
			TangramThreadInfo* pThreadInfo = theApp.GetThreadInfo(dwID);
			theApp.SetHook(dwID);
			map<HWND, CWndFrame*>::iterator iter = pThreadInfo->m_mapTangramFrame.find(m_hWnd);
			if (iter != pThreadInfo->m_mapTangramFrame.end())
			{
				pThreadInfo->m_mapTangramFrame.erase(iter);
				pThreadInfo->m_mapTangramFrame[_hHostWnd] = this;
			}
		}

		m_hHostWnd = UnsubclassWindow(true);
		if(::IsWindow(m_hHostWnd))
		{
			HWND m_hParentWnd = ::GetParent(m_hHostWnd);
			m_Parent.Attach(m_hParentWnd);
			m_Parent.ScreenToClient(&rc);
			m_Parent.Detach();
			::MoveWindow(m_hHostWnd, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, true);
			//::SetWindowPos(m_hHostWnd, HWND_BOTTOM,rc.left,rc.top,rc.right-rc.left,rc.bottom-rc.top, SWP_NOZORDER/*|SWP_FRAMECHANGED|SWP_SHOWWINDOW*/);
		}
	}

	SubclassWindow(_hHostWnd);
	m_hHostWnd = _hHostWnd;
	::GetWindowRect(_hHostWnd,&rc);
	m_Parent.Attach(hParent);
	m_Parent.ScreenToClient(&rc);
	m_Parent.Detach();
	::SetParent(m_pWorkNode->m_pHostWnd->m_hWnd, ::GetParent(_hHostWnd));
	::SetWindowPos(m_pWorkNode->m_pHostWnd->m_hWnd,NULL,rc.left,rc.top,rc.right-rc.left,rc.bottom-rc.top, SWP_FRAMECHANGED|SWP_SHOWWINDOW);
	theApp.HostPosChanged(m_pBindingNode);
	return S_OK;
}

STDMETHODIMP CWndFrame::get_HWND(long* pVal)
{
	*pVal = (long)m_hWnd;
	return S_OK;
}

STDMETHODIMP CWndFrame::get_VisibleNode(IWndNode** pVal)
{
	if (m_pWorkNode)
	{
		*pVal = m_pWorkNode;
	}

	return S_OK;
}

STDMETHODIMP CWndFrame::get_WndPage(IWndPage** pVal)
{
	if (m_pPage)
		*pVal = m_pPage;

	return S_OK;
}

STDMETHODIMP CWndFrame::Extend(BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode)
{
	if (m_pPage)
	{
		DWORD dwID = ::GetWindowThreadProcessId(m_hHostWnd, NULL);
		TRACE(_T("ExtendEx ThreadInfo:%x\n"), dwID);
		TangramThreadInfo* pThreadInfo = theApp.GetThreadInfo(dwID);
		theApp.SetHook(dwID);

		m_strCurrentKey = OLE2T(bstrKey);
		if (m_strCurrentKey == _T(""))
			m_strCurrentKey = _T("Default");
		theApp.m_pPage = m_pPage;
		theApp.m_pWndFrame = this;

		CString strXml = OLE2T(bstrXml);
		theApp.Lock();
		m_strCurrentKey = m_strCurrentKey.MakeLower();
		theApp.m_strCurrentKey = m_strCurrentKey;
		theApp.Unlock();

		auto func = [this](CWndNode* pNode) 
		{
			m_pWorkNode = pNode;
			if (pNode->m_pHostClientView)
			{
				m_pBindingNode = pNode->m_pHostClientView->m_pWndNode;
			}
			else
				m_pBindingNode = nullptr;
		};
		m_pPage->Fire_BeforeExtendXml(CComBSTR(strXml), (long)m_hHostWnd);
		auto it = m_mapNode.find(m_strCurrentKey);
		CWndNode* pOldNode = m_pWorkNode;
		if (it != m_mapNode.end())
		{
			func(it->second);
		}
		else
		{
			CWndNode* pNode = theApp.m_pHostCore->ExtendEx((long)m_hHostWnd, strXml, 0, 0, 0, 0, 0, 0, 0, 0);
			if (pNode == nullptr)
				return S_FALSE;
		}

		if (pOldNode != nullptr&&pOldNode != m_pWorkNode)
		{
			RECT  rc;
			if (pOldNode&&::IsWindow(pOldNode->m_pHostWnd->m_hWnd))
				::GetWindowRect(pOldNode->m_pHostWnd->m_hWnd, &rc);
			CWnd* pWnd = m_pWorkNode->m_pHostWnd;
			HWND hParent = ::GetParent(m_hWnd);
			CWnd::FromHandle(hParent)->ScreenToClient(&rc);
			CWnd* pOldWnd = nullptr;
			if (pOldNode->m_pHostWnd != nullptr&& pWnd != pOldNode->m_pHostWnd&&::IsWindow(pOldNode->m_pHostWnd->m_hWnd))
			{
				pOldWnd = pOldNode->m_pHostWnd;
				::SetWindowLongPtr(pOldNode->m_pHostWnd->GetSafeHwnd(), GWLP_ID, 0);
				::SetParent(pOldWnd->m_hWnd, pWnd->m_hWnd);
			}
			if (pWnd != nullptr&& pWnd != pOldNode->m_pHostWnd)
			{
				::SetWindowLongPtr(pWnd->GetSafeHwnd(), GWLP_ID, m_pWorkNode->m_nID);
				::SetParent(pWnd->m_hWnd, hParent);
				pOldWnd->ShowWindow(SW_HIDE);
			}

			m_pWorkNode->m_bTopObj = true;
			pOldNode->m_bTopObj = false;
			::SetWindowPos(pWnd->m_hWnd, NULL, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, SWP_SHOWWINDOW | SWP_FRAMECHANGED);//|SWP_NOREDRAWSWP_NOZORDER);
		}
		theApp.m_strCurrentKey = _T("");
		if (m_pWorkNode != nullptr)
		{
			if (m_pWorkNode->m_nViewType != Splitter)
			{
				if (m_pWorkNode->m_pHostWnd)
					m_pWorkNode->m_pHostWnd->ModifyStyleEx(WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE, 0);
			}
			theApp.HostPosChanged(m_pWorkNode);
			if (theApp.m_pMDIClientBKWnd&&m_pBKWnd == theApp.m_pMDIClientBKWnd)
				::PostMessage(theApp.m_pMDIClientBKWnd->m_hWnd, WM_NAVIXTML, 0, 0);

			HRESULT hr = S_OK;
			int cConnections = theApp.m_pHostCore->m_vec.GetSize();

			for (int iConnection = 0; iConnection < cConnections; iConnection++)
			{
				CComPtr<IUnknown> punkConnection = theApp.m_pHostCore->m_vec.GetAt(iConnection);
				IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);
				if (pConnection)
				{
					CComVariant avarParams[3];
					avarParams[2] = (long)m_hHostWnd;
					avarParams[2].vt = VT_I4;
					avarParams[1] = strXml.AllocSysString();
					avarParams[1].vt = VT_BSTR;
					avarParams[0] = (IWndNode*)m_pWorkNode;
					CComVariant varResult;
					DISPPARAMS params = { avarParams, NULL, 3, 0 };
					hr = pConnection->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
				}
			}
		}

		*ppRetNode = (IWndNode*)m_pWorkNode;
		return S_OK;
	}

	return S_OK;
}

STDMETHODIMP CWndFrame::get_CurrentNavigateKey(BSTR* pVal)
{
	*pVal = m_strCurrentKey.AllocSysString();
	return S_OK;
}

void CWndFrame::Destroy()
{
	CWndNode* pPage = nullptr;
	CString strPlugID = _T("");
	for (auto it : m_mapNode)
	{
		pPage = it.second;
		if (pPage->m_pTangramParse)
		{
			CTangramXmlParse* pParse = pPage->m_pTangramParse->GetChild(_T("docplugin"));
			if (pParse)
			{
				int nCount = pParse->GetCount();
				for (int i = 0; i < nCount; i++)
				{
					CTangramXmlParse* pChild = pParse->GetChild(i);
					strPlugID = pChild->text();
					strPlugID.Trim();
					if (strPlugID != _T(""))
					{
						if (strPlugID.Find(_T(",")) == -1)
						{
							strPlugID.MakeLower();
							IDispatch* pDisp = (IDispatch*)pPage->m_PlugInDispDictionary[strPlugID];
							if (pDisp)
							{
								pPage->m_PlugInDispDictionary.RemoveKey(LPCTSTR(strPlugID));
								pDisp->Release();
							}
						}
					}
				}
			}

			pPage->m_PlugInDispDictionary.RemoveAll();
		}
	}
}

void CWndFrame::OnFinalMessage(HWND hWnd)
{
	HWND hParent = ::GetParent(hWnd);
	::PostMessage(hParent, WM_TANGRAMMSG, 4096,0);
	__super::OnFinalMessage(hWnd);
}

LRESULT CWndFrame::OnSetHelpWnd(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	switch (wParam)
	{
	case 0:
		m_hLWnd = (HWND)lParam;
		break;
	case 1:
		m_hTWnd = (HWND)lParam;
		break;
	case 2:
		m_hRWnd = (HWND)lParam;
		break;
	case 3:
		m_hBWnd = (HWND)lParam;
		break;
	case 4:
		m_hOuterLWnd = (HWND)lParam;
		break;
	case 5:
		m_hOuterTWnd = (HWND)lParam;
		break;
	case 6:
		m_hOuterRWnd = (HWND)lParam;
		break;
	case 7:
		m_hOuterBWnd = (HWND)lParam;
		break;
	}
	theApp.HostPosChanged(m_pBindingNode);
	return 0;
}

LRESULT CWndFrame::OnSetHelpWndEx(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	switch (wParam)
	{
	case 0:
		m_left2 = (LONG)lParam;
		break;
	case 1:
		m_top2 = (LONG)lParam;
		break;
	case 2:
		m_right2 = (LONG)lParam;
		break;
	case 3:
		m_bottom2 = (LONG)lParam;
		break;
	case 4:
		m_left = (LONG)lParam;
		break;
	case 5:
		m_top = (LONG)lParam;
		break;
	case 6:
		m_right = (LONG)lParam;
		break;
	case 7:
	{
		LONG nOld = m_bottom;
		m_bottom = (LONG)lParam;
		HWND hWnd = ::GetParent(m_pWorkNode->m_pHostWnd->m_hWnd);
		RECT rc;
		m_pWorkNode->m_pHostWnd->GetWindowRect(&rc);
		CWnd::FromHandle(hWnd)->ScreenToClient(&rc);
		rc.bottom -= m_bottom - nOld;
		::SetWindowPos(m_pWorkNode->m_pHostWnd->m_hWnd, HWND_BOTTOM, m_left, m_top, rc.right - rc.left, rc.bottom - rc.top + m_bottom, SWP_NOSENDCHANGING | SWP_NOACTIVATE | SWP_FRAMECHANGED);
		if (::IsWindow(m_hOuterBWnd))
		{
			::SetWindowPos(m_hOuterBWnd, NULL,
				rc.left,
				rc.bottom,
				rc.right - rc.left,
				m_bottom,
				SWP_NOSENDCHANGING | SWP_SHOWWINDOW | SWP_NOACTIVATE | SWP_FRAMECHANGED);
		}
	}
		break;
	}
	theApp.HostPosChanged(m_pBindingNode);
	return 0;
}

LRESULT CWndFrame::OnHScroll(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	LRESULT hr = DefWindowProc(uMsg, wParam, lParam);
	if (m_pBKWnd)
	{
		RECT rect;
		::GetClientRect(m_hWnd, &rect);
		if (::IsWindow(m_pBKWnd->m_hWnd))
			::SetWindowPos(m_pBKWnd->m_hWnd, HWND_BOTTOM, 0, 0, rect.right, rect.bottom, SWP_NOREPOSITION | SWP_NOSENDCHANGING | SWP_NOACTIVATE);
		else
			::InvalidateRect(m_hWnd, &rect, true);
		return hr;
	}
	return hr;
}

LRESULT CWndFrame::OnMouseActivate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	if (m_pBKWnd == nullptr)
	{
		theApp.m_hActiveWnd = 0;
		theApp.m_bWinFormActived = false;
		theApp.m_pWndNode = nullptr;
		theApp.m_pWndFrame = nullptr;
		if (m_bDesignerState)
		{
			theApp.LoadCLR();
			theApp.m_pCloudAddinCLRProxy->SelectNode(m_pBindingNode);
		}
	}
	return DefWindowProc(uMsg, wParam, lParam);
}

LRESULT CWndFrame::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	if (theApp.m_pDesigningFrame == this)
		theApp.m_pDesigningFrame = nullptr;
	m_pPage->BeforeDestory();
	auto it = theApp.m_mapWindowPage.find(m_pPage->m_hWnd);
	if (it != theApp.m_mapWindowPage.end())
		theApp.m_mapWindowPage.erase(it);

	return DefWindowProc(uMsg, wParam, lParam);
}

LRESULT CWndFrame::OnTangramMsg(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	if (lParam == 2048)
	{
		CtrlInfo* pCtrlInfo = (CtrlInfo*)wParam;
		if (pCtrlInfo&&pCtrlInfo->m_pPage)
		{
			CWndPage* pPage = (CWndPage*)pCtrlInfo->m_pPage;
			pPage->Fire_ClrControlCreated(pCtrlInfo->m_pNode, pCtrlInfo->m_pCtrlDisp, pCtrlInfo->m_strName.AllocSysString(), (long)pCtrlInfo->m_hWnd);
		}
	}
	return DefWindowProc(uMsg, wParam, lParam);
}

LRESULT CWndFrame::OnGetMe(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	if (lParam == 1992)
	{
		return (LRESULT)this;
	}

	return DefWindowProc(uMsg, wParam, lParam);
}

LRESULT CWndFrame::OnNcDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	LONG_PTR pfnWndProc = ::GetWindowLongPtr(m_hWnd, GWLP_WNDPROC);
	LRESULT lr = DefWindowProc(uMsg, wParam, lParam);
	if (m_pfnSuperWindowProc != ::DefWindowProc && ::GetWindowLongPtr(m_hWnd, GWLP_WNDPROC) == pfnWndProc)
		::SetWindowLongPtr(m_hWnd, GWLP_WNDPROC, (LONG_PTR)m_pfnSuperWindowProc);

	// mark window as destryed
	m_dwState |= WINSTATE_DESTROYED;

	return lr;
}

LRESULT CWndFrame::OnWindowPosChanging(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
	WINDOWPOS* lpwndpos = (WINDOWPOS*)lParam;
	LRESULT hr = S_OK;
	if(m_pWorkNode)
	{
		if(m_bNeedOuterChanged)
		{
			::SetWindowPos(m_pWorkNode->m_pHostWnd->m_hWnd, HWND_BOTTOM, lpwndpos->x + m_left, lpwndpos->y + m_top, lpwndpos->cx - m_right, lpwndpos->cy - m_bottom, lpwndpos->flags | SWP_NOSENDCHANGING | SWP_NOREDRAW | SWP_NOACTIVATE | SWP_FRAMECHANGED);
			if (::IsWindow(m_hOuterBWnd))
			{
				::SetWindowPos(m_hOuterBWnd, NULL,
					lpwndpos->x + m_left,
					lpwndpos->y + m_top + lpwndpos->cy - m_bottom,
					lpwndpos->cx - m_right,
					m_bottom,
					lpwndpos->flags | SWP_NOSENDCHANGING | SWP_SHOWWINDOW | SWP_NOACTIVATE | SWP_FRAMECHANGED);
			}
		}
		else
		{
			if (m_bOuterInited == false && lpwndpos->cy&&::IsWindow(m_hOuterBWnd))
			{
				m_bOuterInited = true;
				::SetWindowPos(m_pWorkNode->m_pHostWnd->m_hWnd, HWND_BOTTOM, lpwndpos->x + m_left, lpwndpos->y + m_top, lpwndpos->cx - m_right, lpwndpos->cy - m_bottom, lpwndpos->flags | SWP_NOSENDCHANGING | SWP_NOACTIVATE | SWP_FRAMECHANGED);
				if (::IsWindow(m_hOuterBWnd))
				{
					::SetWindowPos(m_hOuterBWnd, NULL,
						lpwndpos->x + m_left,
						lpwndpos->y + m_top + lpwndpos->cy - m_bottom,
						lpwndpos->cx - m_right,
						m_bottom,
						lpwndpos->flags | SWP_NOSENDCHANGING | SWP_SHOWWINDOW | SWP_NOACTIVATE | SWP_FRAMECHANGED);
				}
			}
			else
				::SetWindowPos(m_pWorkNode->m_pHostWnd->m_hWnd, HWND_BOTTOM, lpwndpos->x, lpwndpos->y, lpwndpos->cx, lpwndpos->cy, lpwndpos->flags | SWP_NOSENDCHANGING | SWP_NOACTIVATE | SWP_FRAMECHANGED);
			m_bNeedOuterChanged = true;
		}
		if (m_pBindingNode != nullptr&&::IsWindow(m_pBindingNode->m_pHostWnd->m_hWnd))
		{
			CRect rect, rect2;
			CWnd* pWnd = m_pBindingNode->m_pHostWnd;

			if (::IsWindow(pWnd->m_hWnd))
			{
				pWnd->GetWindowRect(&rect);

				::IsWindowVisible(pWnd->m_hWnd) ? lpwndpos->flags |= SWP_SHOWWINDOW : lpwndpos->flags |= SWP_HIDEWINDOW;
				CWnd::FromHandle(::GetParent(m_hWnd))->ScreenToClient(&rect);
				lpwndpos->x = rect.left + m_left2;
				if (::IsWindow(m_hLWnd))
				{
					rect2.left = rect.left;
					rect2.right = rect.left + m_left2;
					rect2.top = rect.top;
					rect2.bottom = rect.top + rect.Height() - m_bottom2;
					::SetWindowPos(m_hLWnd, HWND_TOP, rect2.left, rect2.top, rect2.Width(), rect2.Height(), lpwndpos->flags | SWP_NOSENDCHANGING | SWP_NOACTIVATE | SWP_FRAMECHANGED);
				}
				lpwndpos->y = rect.top + m_top2;
				if (::IsWindow(m_hTWnd))
				{
					rect2.left = 0;
					rect2.top = 0;
					rect2.right = lpwndpos->cx;
					rect2.bottom = m_top2;
					::SetWindowPos(m_hTWnd, NULL, rect2.left, rect2.top, rect2.Width(), rect2.Height(), lpwndpos->flags | SWP_NOSENDCHANGING | SWP_NOACTIVATE | SWP_FRAMECHANGED);
				}
				lpwndpos->cx = rect.Width() - m_right2 - m_left2;
				if (::IsWindow(m_hRWnd))
				{
					rect2.left = 0;
					rect2.top = 0;
					rect2.right = lpwndpos->cx;
					rect2.bottom = m_top2;
					::SetWindowPos(m_hRWnd, NULL, rect2.left, rect2.top, rect2.Width(), rect2.Height(), lpwndpos->flags | SWP_NOSENDCHANGING | SWP_NOACTIVATE | SWP_FRAMECHANGED);
				}
				lpwndpos->cy = rect.Height() - m_bottom2;
				if (::IsWindow(m_hBWnd))
				{
					rect2.left = lpwndpos->x - m_left2;
					rect2.top = lpwndpos->y + lpwndpos->cy;
					rect2.right = lpwndpos->x + lpwndpos->cx;
					rect2.bottom = lpwndpos->y + lpwndpos->cy + m_bottom2;
					::SetWindowPos(m_hBWnd, NULL, rect2.left, rect2.top, rect2.Width(), rect2.Height(), lpwndpos->flags | SWP_NOSENDCHANGING | SWP_NOACTIVATE | SWP_FRAMECHANGED);
				}
				if (::IsWindow(m_hOuterBWnd) && ::IsWindowVisible(m_hOuterBWnd) == false)
					::ShowWindow(m_hOuterBWnd, SW_SHOW);
			}
		}
		else
			lpwndpos->flags |= SWP_HIDEWINDOW;
	}

	return lRes;
}

LRESULT CWndFrame::OnParentNotify(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	theApp.m_pWndFrame = nullptr;
	return DefWindowProc(uMsg, wParam, lParam);
}

//LRESULT CWndFrame::OnEraseBkgnd(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
//{
//	HDC hDC = (HDC)wParam;
//	CDC dc;
//	dc.Attach(hDC);
//	RECT rt;
//	CBitmap bit;
//	GetClientRect(&rt);
//	bit.LoadBitmap(IDB_BITMAP_GRID);
//	CBrush br(&bit);
//	dc.FillRect(&rt, &br);
//	dc.SetBkMode(TRANSPARENT);
//	CString strInfo = _T("");
//	strInfo = strInfo +
//		_T("\n  ----Tangram Object Information----") +
//		_T("\n  ") +
//		_T("\n   Object Name:   ") + _T("") +
//		_T("\n   Object Caption:") + 
//		_T("\n  ") +
//		_T("\n  ");
//
//	dc.SetTextColor(RGB(255, 255, 255));
//	dc.DrawText(strInfo, &rt, DT_LEFT);
//	dc.Detach();
//
//	return true;
//}

STDMETHODIMP CWndFrame::get_DesignerState(VARIANT_BOOL* pVal)
{
	if (m_bDesignerState)
		*pVal = true;
	else
		*pVal = false;

	return S_OK;
}

STDMETHODIMP CWndFrame::put_DesignerState(VARIANT_BOOL newVal)
{
	m_bDesignerState = newVal;
	return S_OK;
}

STDMETHODIMP CWndFrame::GetXml(BSTR bstrRootName, BSTR* bstrRet)
{
	CString strRootName = OLE2T(bstrRootName);
	if (strRootName == _T(""))
		strRootName = _T("DocumentUI");
	CString strXmlData = _T("<Default><window><node name=\"Start\"/></window></Default>");
	CString strName = _T("");
	CString strXml = _T("");

	map<CString, CString> m_mapTemp;
	map<CString, CString>::iterator it2;
	for (auto it : m_mapNode)
	{
		theApp.m_pHostCore->UpdateWndNode(it.second);
		strName = it.first;
		int nPos = strName.Find(_T("-"));
		CString str = strName.Mid(nPos+1);
		if (str.CompareNoCase(_T("inDesigning")) == 0)
		{
			strName = strName.Left(nPos);
			m_mapTemp[strName] = it.second->m_pTangramParse->xml();
		}
	}

	for (auto it : m_mapNode)
	{
		strName = it.first;
		if (strName.Find(_T("-indesigning")) == -1)
		{
			it2 = m_mapTemp.find(strName);
			if (it2 != m_mapTemp.end())
				strXml = it2->second;
			else
				strXml = it.second->m_pTangramParse->xml();
			strXmlData += strXml;
		}
	}
	
	strXml = _T("<");
	strXml += strRootName;
	strXml += _T(">");
	strXml += strXmlData;
	strXml += _T("</");
	strXml += strRootName;
	strXml += _T(">");
	*bstrRet = strXml.AllocSysString();
	return S_OK;
}

