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

// TangramObject.cpp : Implementation of CWndPage

#include "stdafx.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "CloudAddinApp.h"
#include "CloudAddinCore.h"

// CWndPage

CWndPage::CWndPage()
{
	m_hWnd								= 0;
	m_nWebViewCount						= 0;
	m_nThreadId							= GetCurrentThreadId();
	m_bIsBlank							= false;
	m_bIsDestory						= false;
	m_bDocComplete						= false;
	m_bEnableSinkCLRCtrlCreatedEvent	= true;
	m_strURL							= _T("");
	m_pCP								= nullptr;
	m_pExternalDisp						= nullptr;
	m_pHtmlDocument2					= nullptr;
	m_pHTMLWindow2						= nullptr;
	theApp.m_pPage						= this;
#ifdef _DEBUG
	theApp.m_nTangram++;
#endif	
}

CWndPage::~CWndPage()
{
#ifdef _DEBUG
	theApp.m_nTangram--;
#endif	
	if (theApp.m_mapWindowPage.size() == 0)
		theApp.Close();

	UnLoad();

	if (m_hWnd == theApp.m_pHostCore->m_hMainWnd)
	{
		auto it2 = theApp.m_pHostCore->m_mapObjDic.find(_T("defaulttangram"));
		if (it2 != theApp.m_pHostCore->m_mapObjDic.end())
			theApp.m_pHostCore->m_mapObjDic.erase(it2);
		it2 = theApp.m_pHostCore->m_mapObjDic.find(_T("defaultframe"));
		if (it2 != theApp.m_pHostCore->m_mapObjDic.end())
			theApp.m_pHostCore->m_mapObjDic.erase(it2);
	}

	for (auto it2 : m_mapExternalObj)
	{
		it2.second->Release();
	}
	m_mapExternalObj.erase(m_mapExternalObj.begin(), m_mapExternalObj.end());

	m_mapFrame.erase(m_mapFrame.begin(), m_mapFrame.end());
	m_mapNode.erase(m_mapNode.begin(), m_mapNode.end());
	auto it = theApp.m_mapWindowPage.find(m_hWnd);
	if (it != theApp.m_mapWindowPage.end())
	{
		theApp.m_mapWindowPage.erase(it);
	}
}

STDMETHODIMP CWndPage::get_Count(long* pCount)
{
	*pCount = (long)m_mapFrame.size();
	return S_OK;
}

STDMETHODIMP CWndPage::get_Frame(VARIANT vIndex, IWndFrame **ppFrame)
{
	if (vIndex.vt == VT_I4)
	{
		long lCount = m_mapFrame.size();
		int iIndex = vIndex.lVal;
		if (iIndex < 0 || iIndex >= lCount) return E_INVALIDARG;

		map<HWND, CWndFrame*>::iterator it;
		it = m_mapFrame.begin();
		int iPos = 0;
		while (it != m_mapFrame.end())
		{
			if (iPos == iIndex) break;
			iPos++;
			it++;
		}

		CWndFrame* pFrame = it->second;
		*ppFrame = pFrame;
		return S_OK;
	}

	if (vIndex.vt == VT_BSTR)
	{
		CString strKey = OLE2T(vIndex.bstrVal);
		auto it = m_mapWnd.find(strKey);
		if (it != m_mapWnd.end())
		{
			auto it2 = m_mapFrame.find(it->second);
			if (it2 != m_mapFrame.end())
			{
				CWndFrame* pFrame = it2->second;
				*ppFrame = pFrame;
				return S_OK;
			}
		}
		return E_INVALIDARG;
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get__NewEnum(IUnknown** ppVal)
{
	typedef CComEnum<IEnumVARIANT, &IID_IEnumVARIANT, VARIANT, _Copy<VARIANT>>
		CComEnumVariant;

	CComObject<CComEnumVariant> *pe = 0;
	HRESULT hr = pe->CreateInstance(&pe);

	if (SUCCEEDED(hr))
	{
		pe->AddRef();
		long nLen = (long)m_mapFrame.size();
		VARIANT* rgvar = new VARIANT[nLen];
		ZeroMemory(rgvar, sizeof(VARIANT)*nLen);
		VARIANT* pItem = rgvar;
		for (auto it : m_mapFrame)
		{
			IUnknown* pDisp = nullptr;
			CWndFrame* pObj = it.second;
			hr = pObj->QueryInterface(IID_IUnknown, (void**)&pDisp);
			pItem->vt = VT_UNKNOWN;
			pItem->punkVal = pDisp;
			if (pItem->punkVal != nullptr)
				pItem->punkVal->AddRef();
			pItem++;
			pDisp->Release();
		}

		hr = pe->Init(rgvar, &rgvar[nLen], 0, AtlFlagTakeOwnership);
		if (SUCCEEDED(hr))
			hr = pe->QueryInterface(IID_IUnknown, (void**)ppVal);
		pe->Release();
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_External(IDispatch** ppVal)
{
	if (m_pExternalDisp)
	{
		*ppVal = m_pExternalDisp;
		(*ppVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndPage::put_External(IDispatch* newVal)
{
	if (m_pExternalDisp)
	{
		UnSinkObject(m_pExternalDisp, nullptr);
		m_pExternalDisp->Release();
	}
	if (newVal == nullptr)
	{
		m_pExternalDisp = nullptr;
		return S_OK;
	}
	m_pExternalDisp = newVal;
	m_pExternalDisp->AddRef();
	return S_OK;
}

STDMETHODIMP CWndPage::CreateFrame(VARIANT ParentObj, VARIANT HostWnd, BSTR bstrFrameName, IWndFrame** pRetFrame)
{
	HWND h = 0; 
	if (ParentObj.vt == VT_DISPATCH&&HostWnd.vt == VT_BSTR)
	{
		if (theApp.m_pCloudAddinCLRProxy)
		{
			IDispatch* pDisp = nullptr;
			pDisp = theApp.m_pCloudAddinCLRProxy->GetCtrlByName(ParentObj.pdispVal, HostWnd.bstrVal, true);
			if (pDisp)
			{
				h = theApp.m_pCloudAddinCLRProxy->GetCtrlHandle(pDisp);
				if (h)
				{
					return CreateFrame(CComVariant(0), CComVariant((long)h), bstrFrameName, pRetFrame);
				}
			}
		}
		return S_FALSE;
	}
	if (HostWnd.vt == VT_DISPATCH)
	{
		if (theApp.m_pCloudAddinCLRProxy)
		{
			h=theApp.m_pCloudAddinCLRProxy->IsCtrlCanNavigate(HostWnd.pdispVal);
			if (h)
			{
				CString strName = OLE2T(bstrFrameName);
				if (strName == _T(""))
				{
					bstrFrameName = theApp.m_pCloudAddinCLRProxy->GetCtrlName(HostWnd.pdispVal);
					strName = OLE2T(bstrFrameName);
					if (strName == _T(""))
						bstrFrameName = CComBSTR(L"Frame");
				}
				auto it = m_mapFrame.find((HWND)h);
				if (it == m_mapFrame.end())
					return CreateFrame(CComVariant(0), CComVariant((long)h), bstrFrameName, pRetFrame);
				else
					*pRetFrame = it->second;
			}
		}
	}
	else if (HostWnd.vt == VT_I2||HostWnd.vt == VT_I4 || HostWnd.vt == VT_I8)
	{
		HWND _hWnd = NULL;
		if(HostWnd.vt == VT_I4)
			_hWnd = (HWND)HostWnd.lVal;
		if(HostWnd.vt == VT_I8)
			_hWnd = (HWND)HostWnd.llVal;
		if (_hWnd == 0)
		{
			_hWnd = ::FindWindowEx(m_hWnd, NULL, _T("MDIClient"), NULL);
			if(_hWnd==NULL)
				_hWnd = ::GetWindow(m_hWnd, GW_CHILD);
		}
		if (_hWnd&&::IsWindow(_hWnd))
		{
			map<HWND, CWndFrame*>::iterator it = m_mapFrame.find(_hWnd);
			if (it == m_mapFrame.end())
			{
				DWORD dwID = ::GetWindowThreadProcessId(_hWnd, NULL);
				TangramThreadInfo* pThreadInfo = theApp.GetThreadInfo(dwID);

				CWndFrame* m_pExtenderFrame = new CComObject<CWndFrame>();
				if (_hWnd == theApp.m_pHostCore->m_hMainFrameWnd)
				{
					theApp.m_pHostCore->m_mapObjDic[_T("defaultframe")] = m_pExtenderFrame;
				}
				CString strName = OLE2T(bstrFrameName);
				if (strName == _T(""))
					strName = _T("Frame");
				m_pExtenderFrame->m_strFrameName = strName;
				m_pExtenderFrame->m_pPage = this;
				m_pExtenderFrame->m_hHostWnd = _hWnd;
				pThreadInfo->m_mapTangramFrame[_hWnd] = m_pExtenderFrame;
				m_mapFrame[_hWnd] = m_pExtenderFrame;
				m_mapWnd[strName] = _hWnd;
				*pRetFrame = m_pExtenderFrame;
			}
			else
				*pRetFrame = it->second;
		}
	}
		
	return S_OK;
}


STDMETHODIMP CWndPage::GetFrameFromCtrl(IDispatch* CtrlObj, IWndFrame** ppFrame)
{
	if (theApp.m_pCloudAddinCLRProxy)
	{
		HWND h = theApp.m_pCloudAddinCLRProxy->IsCtrlCanNavigate(CtrlObj);
		if (h)
		{
			auto it = m_mapFrame.find(h);
			if (it != m_mapFrame.end())
				*ppFrame = it->second;
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::GetWndNode(BSTR bstrFrameName, BSTR bstrNodeName, IWndNode** pRetNode)
{
	CString strKey = OLE2T(bstrFrameName);
	CString strName = OLE2T(bstrNodeName);
	if (strKey == _T("") || strName == _T(""))
		return S_FALSE;
	auto it = m_mapWnd.find(strKey);
	if (it != m_mapWnd.end())
	{
		HWND hWnd = it->second;
		if (::IsWindow(hWnd))
		{
			auto it = m_mapFrame.find(hWnd);
			if (it != m_mapFrame.end())
			{
				CWndFrame* pFrame = it->second;
				strName = strName.MakeLower();
				auto it2 = pFrame->m_mapNode.find(strName);
				if (it2 != pFrame->m_mapNode.end())
					*pRetNode = (IWndNode*)it2->second;
				else
				{
					it2 = m_mapNode.find(strName);
					if (it2 != m_mapNode.end())
						*pRetNode = (IWndNode*)it2->second;
				}
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_URL(BSTR* pVal)
{
	*pVal = m_strURL.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndPage::put_URL(BSTR newVal)
{
	m_strURL = OLE2T(newVal);
	UnLoad();
	Init(m_strURL);
	return S_OK;
}

STDMETHODIMP CWndPage::AddObjToHtml(BSTR strObjName, VARIANT_BOOL bConnectEvent, IDispatch* pObjDisp)
{
	CString strName = OLE2T(strObjName);
	if (strName!=_T(""))
	{
		CJSExtender::AddObject(strName, pObjDisp);
		if (bConnectEvent)
		{
			CComQIPtr<IConnectionPointContainer> pContainer(pObjDisp);
			if (pContainer)
			{
				strName += _T("_On");
				SinkObject(strName, pObjDisp);
			}
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_Extender(BSTR bstrExtenderName, IDispatch** pVal)
{
	CString strName = OLE2T(bstrExtenderName);
	if (strName == _T(""))
		return S_OK;
	auto it = m_mapExternalObj.find(strName);
	if (it != m_mapExternalObj.end())
	{
		*pVal = it->second;
		(*pVal)->AddRef();
	}
	return S_OK;
}


STDMETHODIMP CWndPage::put_Extender(BSTR bstrExtenderName, IDispatch* newVal)
{
	CString strName = OLE2T(bstrExtenderName);
	if (strName == _T(""))
		return S_OK;
	auto it = m_mapExternalObj.find(strName);
	if (it == m_mapExternalObj.end())
	{
		m_mapExternalObj[strName] = newVal;
		newVal->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_FrameName(long hHwnd, BSTR* pVal)
{
	HWND _hWnd = (HWND)hHwnd;
	if (_hWnd)
	{
		auto it = m_mapFrame.find(_hWnd);
		if (it!=m_mapFrame.end())
			*pVal = it->second->m_strFrameName.AllocSysString();
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_WndFrame(BSTR bstrFrameName, IWndFrame** pVal)
{
	CString strName = OLE2T(bstrFrameName);
	if (strName != _T(""))
	{
		map<CString, HWND>::iterator it = m_mapWnd.find(strName);
		if (it != m_mapWnd.end())
		{
			HWND h = it->second;
			map<HWND, CWndFrame*>::iterator it2 = m_mapFrame.find(h);
			if (it2 != m_mapFrame.end())
				*pVal = it2->second;
		}
	}
	return S_OK;
}

void CWndPage::OnNodeDocComplete(WPARAM wParam)
{
	bool bState = false;
	for (auto it1 : m_mapFrame)
	{
		for (auto it2 : it1.second->m_mapNode)
		{
			if (it2.second->m_bCreated == false)
			{
				::PostMessage(m_hWnd, WM_TANGRAM_WEBNODEDOCCOMPLETE, wParam, 0);
				return;
			}
		}
	}

	switch (wParam)
	{
	case 0:
		{
			int nIndex = 0;
			for (auto it : m_mapNode)
			{
				if (it.second->m_bNodeDocComplete==false)
				{
					it.second->m_bNodeDocComplete = true;
					BSTR bstrURL = it.second->m_strURL.AllocSysString();
					it.second->Fire_NodeDocumentComplete(it.second->m_pExtender, bstrURL);
					CComPtr<IHTMLWindow2> pWinDisp;
					it.second->m_pTangramNodeWBEvent->m_pHTMLDocument2->get_parentWindow(&pWinDisp);
					Fire_PageLoaded(pWinDisp.p, bstrURL);
					::SysFreeString(bstrURL);
					nIndex++;
				}
			}
			if (nIndex < m_nWebViewCount)
				::PostMessage(m_hWnd, WM_TANGRAM_WEBNODEDOCCOMPLETE, wParam, 0);
		}
		break;
	case 1:
		if (m_bDocComplete == false)
		{
			m_bDocComplete = true;
			CComQIPtr<IHTMLDocument3> pDoc(m_pHtmlDocument2);
			CComPtr<IHTMLElementCollection> pColl;
			HRESULT hr = pDoc->getElementsByTagName(CComBSTR(L"xml"), &pColl);
			if (hr == S_OK)
			{
				long nCount = 0;
				pColl->get_length(&nCount);
				if (nCount)
				{
					CComBSTR bstrID(L"");
					CComBSTR bstrXml(L"");
					CString strID = _T("");
					CString strXml = _T("");
					int nIndex = 0;
					for (int i = 0; i < nCount; i++)
					{
						CComVariant vIdx((long)i, VT_I4);
						CComPtr<IDispatch> pDisp;
						hr = pColl->item(vIdx, vIdx, &pDisp);
						if (pDisp)
						{
							CComQIPtr<IHTMLElement> pElem(pDisp);
							pElem->get_innerHTML(&bstrXml);
							strXml = OLE2T(bstrXml);
							if (strXml != _T(""))
							{
								pElem->get_id(&bstrID);
								strID = OLE2T(bstrID);
								if (strID == _T(""))
								{
									strID.Format(_T("xtml%d"), nIndex);
									nIndex++;
								}
								m_mapXtml[strID] = strXml;
							}
						}
					}
				}
			}

			Fire_PageLoaded(m_pHTMLWindow2, CComBSTR(m_strURL));
		}
		break;
	}
}

void CWndPage::BeforeDestory()
{
	if (!m_bIsDestory)
	{
		m_bIsDestory = true;
		Fire_Destroy();
		for (auto it : m_mapNode)
		{
			if (it.second->m_pTangramNodeWBEvent)
			{
				it.second->m_pTangramNodeWBEvent->ConnectJSEngine(nullptr);
				it.second->m_pTangramNodeWBEvent->DispEventUnadvise(it.second->m_pDisp);
				delete it.second->m_pTangramNodeWBEvent;
				it.second->m_pTangramNodeWBEvent = nullptr;
			}
		}

		for (auto it: m_mapFrame)
			it.second->Destroy();

		ConnectJSEngine(nullptr);
	}
}

STDMETHODIMP CWndPage::get_Nodes(BSTR bstrNodeName, IWndNode** pVal)
{
	CString strName = OLE2T(bstrNodeName);
	if (strName == _T(""))
		return S_OK;
	auto it2 = m_mapNode.find(strName);
	if (it2 == m_mapNode.end())
		return S_OK;

	if (it2->second)
		*pVal = it2->second;

	return S_OK;
}

STDMETHODIMP CWndPage::get_XObjects(BSTR bstrName, IDispatch** pVal)
{
	CString strName = OLE2T(bstrName);
	if (strName == _T(""))
		return S_OK;
	auto it2 = m_mapNode.find(strName);
	if (it2 == m_mapNode.end())
		return S_OK;
	if (it2->second->m_pDisp)
	{
		*pVal = it2->second->m_pDisp;
		(*pVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndPage::AttachEvent(BSTR bstrName, IDispatch* pWndDisp, VARIANT_BOOL* bResult)
{
	CString strNames = OLE2T(bstrName);
	if (strNames==_T(""))
		return S_FALSE;
	CComQIPtr<IHTMLWindow2> pWindow2(pWndDisp);
	if (pWindow2 == nullptr)
		return S_FALSE;

	CJSExtender* _pJSExtender = nullptr;
	for (auto it : m_mapJSExtender)
	{
		CComPtr<IHTMLWindow2> pWindow;
		it.first->get_parentWindow(&pWindow);
		if (pWindow.p == pWindow2.p)
		{
			_pJSExtender = it.second;
			break;
		}
	}
	if (_pJSExtender == nullptr)
		return S_FALSE;

	int nPos = strNames.Find(_T(","));
	if (nPos == -1)
	{
		auto it = m_mapNode.find(strNames);
		if (it == m_mapNode.end())
		{
			return S_OK;
		}
		strNames += _T("_On");
		_pJSExtender->SinkObject(strNames, it->second->m_pDisp);
		*bResult = true;
	}
	while (nPos != -1)
	{
		CString strName = strNames.Left(nPos);
		strNames = strNames.Mid(nPos + 1);
		nPos = strNames.Find(_T(","));
		auto it = m_mapNode.find(strName);
		if (it != m_mapNode.end() && it->second->m_pDisp)
		{

			strName += _T("_On");
			_pJSExtender->SinkObject(strName, it->second->m_pDisp);
			*bResult = true;
		}
		if (nPos == -1)
		{
			it = m_mapNode.find(strNames);
			if (it != m_mapNode.end())
			{
				strNames += _T("_On");
				_pJSExtender->SinkObject(strNames, it->second->m_pDisp);
				*bResult = true;
				return S_OK;
			}
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_Width(long* pVal)
{
	if (m_hWnd)
	{
		RECT rc;
		::GetWindowRect(m_hWnd, &rc);
		*pVal = rc.right - rc.left;
	}

	return S_OK;
}

STDMETHODIMP CWndPage::put_Width(long newVal)
{
	if (m_hWnd&&newVal)
	{
		RECT rc;
		::GetWindowRect(m_hWnd, &rc);
		rc.right = rc.left + newVal;
		::SetWindowPos(m_hWnd, NULL, rc.left, rc.top, newVal, rc.bottom - rc.top, SWP_NOACTIVATE | SWP_NOREDRAW);
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_Height(long* pVal)
{
	if (m_hWnd)
	{
		RECT rc;
		::GetWindowRect(m_hWnd, &rc);
		*pVal = rc.bottom - rc.top;
	}
	return S_OK;
}

STDMETHODIMP CWndPage::put_Height(long newVal)
{
	if (m_hWnd&&newVal)
	{
		RECT rc;
		::GetWindowRect(m_hWnd, &rc);
		rc.right = rc.left + newVal;
		::SetWindowPos(m_hWnd, NULL, rc.left, rc.top, rc.right - rc.left, newVal, SWP_NOACTIVATE | SWP_NOREDRAW);
	}
	return S_OK; 
}

STDMETHODIMP CWndPage::AddNodesToPage(IDispatch* pHtmlDoc, BSTR bstrNodeNames, VARIANT_BOOL bAdd, VARIANT_BOOL* bSuccess)
{
	CString strNames = OLE2T(bstrNodeNames);
	if (strNames != _T(""))
	{
		map<CString, CWndNode*>::iterator it;
		int nPos = strNames.Find(_T(","));
		if (nPos == -1)
		{
			it = m_mapNode.find(strNames);
			if (it != m_mapNode.end() && it->second->m_pDisp)
			{
				AddObject(pHtmlDoc, it->second->m_pDisp, bstrNodeNames, false, bSuccess);
			}
			return S_OK;
		}
		while (nPos != -1)
		{
			CString strName = strNames.Left(nPos);
			strNames = strNames.Mid(nPos + 1);
			nPos = strNames.Find(_T(","));
			it = m_mapNode.find(strName);
			if (it != m_mapNode.end() && it->second->m_pDisp)
			{
				AddObject(pHtmlDoc, it->second->m_pDisp, strName.AllocSysString(), false, bSuccess);
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_NodeNames(BSTR* pVal)
{
	CString strNames = _T("");
	for (auto it : m_mapNode)
	{
		strNames += it.first;
		strNames += _T(",");
	}
	int nPos = strNames.ReverseFind(',');
	strNames = strNames.Left(nPos);
	*pVal = strNames.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndPage::AddObject(IDispatch* pHtmlDoc, IDispatch* pObject, BSTR bstrObjName, VARIANT_BOOL bSinkEvent, VARIANT_BOOL* bResult)
{
	CString strName = OLE2T(bstrObjName);
	CComQIPtr<IHTMLDocument2> pDoc(pHtmlDoc);
	if (strName != _T("") && pHtmlDoc)
	{
		auto it = m_mapJSExtender.find(pDoc.p);
		if (it != m_mapJSExtender.end())
		{
			CJSExtender* _pJSExtender = it->second;
			if (_pJSExtender)
			{
				_pJSExtender->AddObject(strName, pObject);
				if (bSinkEvent)
				{
					CComQIPtr<IConnectionPointContainer> pContainer(pObject);
					if (pContainer)
					{
						strName += _T("_On");
						_pJSExtender->SinkObject(strName, pObject);
						*bResult = true;
					}
				}
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_HTMLWindow(BSTR NodeName, IDispatch** pVal)
{
	CString strName = OLE2T(NodeName);
	if (strName != _T(""))
	{
		auto it = m_mapNode.find(strName);
		if (it != m_mapNode.end() && it->second->m_pTangramNodeWBEvent->m_pHTMLDocument2)
		{
			CComPtr<IHTMLWindow2> pHTMLWindow2;
			it->second->m_pTangramNodeWBEvent->m_pHTMLDocument2->get_parentWindow(&pHTMLWindow2);
			*pVal = pHTMLWindow2.Detach();
			(*pVal)->AddRef();
		}
	}
	else if(m_pHtmlDocument2)
	{
		CComPtr<IHTMLWindow2> pHTMLWindow2;
		m_pHtmlDocument2->get_parentWindow(&pHTMLWindow2);
		*pVal = pHTMLWindow2.Detach();
		(*pVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_HtmlDocument(IDispatch** pVal)
{
	if (m_pHtmlDocument2)
	{
		*pVal = m_pHtmlDocument2;
		(*pVal)->AddRef();
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_Parent(IWndPage** pVal)
{
	HWND hWnd = ::GetParent(m_hWnd);
	if (hWnd == NULL)
		return S_OK;

	auto it = theApp.m_mapWindowPage.find(hWnd);
	while (it == theApp.m_mapWindowPage.end())
	{
		hWnd = ::GetParent(hWnd);
		if (hWnd == NULL)
			return S_OK;
		it = theApp.m_mapWindowPage.find(hWnd);
		if (it != theApp.m_mapWindowPage.end())
		{
			*pVal = it->second;
			return S_OK;
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::GetHtmlDocument(IDispatch* HtmlWindow, IDispatch** ppHtmlDoc)
{
	CComQIPtr<IHTMLWindow2> pWindow2(HtmlWindow);
	if (pWindow2 == nullptr)
		return S_OK;

	for (auto it : m_mapJSExtender)
	{
		CComPtr<IHTMLWindow2> pWindow;
		it.first->get_parentWindow(&pWindow);
		if (pWindow.p == pWindow2.p)
		{
			*ppHtmlDoc = it.first;
			(*ppHtmlDoc)->AddRef();
			return S_OK;
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::AttachNodeEvent(BSTR bstrNames, IDispatch* pWndDisp)
{
	CString strNames = OLE2T(bstrNames);
	if (strNames == _T(""))
		return S_FALSE;
	CComQIPtr<IHTMLWindow2> pWindow2(pWndDisp);
	if (pWindow2 == NULL)
		return S_FALSE;

	CJSExtender* _pJSExtender = nullptr;
	for (auto it0 : m_mapJSExtender)
	{
		CComPtr<IHTMLWindow2> pWindow;
		it0.first->get_parentWindow(&pWindow);
		if (pWindow.p == pWindow2.p)
		{
			_pJSExtender = it0.second;
			break;
		}
	}
	if (_pJSExtender == nullptr)
		return S_FALSE;

	map<CString, CWndNode*>::iterator it;
	int nPos = strNames.Find(_T(","));
	if (nPos == -1)
	{
		it = m_mapNode.find(strNames);
		if (it == m_mapNode.end())
		{
			return S_OK;
		}
		strNames += _T("_Node_On");
		_pJSExtender->SinkObject(strNames, (IWndNode*)it->second);
	}
	while (nPos != -1)
	{
		CString strName = strNames.Left(nPos);
		strNames = strNames.Mid(nPos + 1);
		nPos = strNames.Find(_T(","));
		it = m_mapNode.find(strName);
		if (it != m_mapNode.end() && it->second->m_pDisp)
		{
			strName += _T("_Node_On");
			_pJSExtender->SinkObject(strName, (IWndNode*)it->second);
		}
		if (nPos == -1)
		{
			it = m_mapNode.find(strNames);
			if (it != m_mapNode.end())
			{
				strNames += _T("_Node_On");
				_pJSExtender->SinkObject(strNames, (IWndNode*)it->second);
				return S_OK;
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::GetCtrlByName(IDispatch* pCtrl, BSTR bstrName, VARIANT_BOOL bFindInChild, IDispatch** ppRetDisp)
{
	if (theApp.m_pCloudAddinCLRProxy)
	{
		*ppRetDisp = theApp.m_pCloudAddinCLRProxy->GetCtrlByName(pCtrl, bstrName, bFindInChild?true:false);
	}
	return S_OK;
}

STDMETHODIMP CWndPage::GetCtrlValueByName(IDispatch* pCtrl, BSTR bstrName, VARIANT_BOOL bFindInChild, BSTR* bstrVal)
{
	if (theApp.m_pCloudAddinCLRProxy)
	{
		*bstrVal = theApp.m_pCloudAddinCLRProxy->GetCtrlValueByName(pCtrl, bstrName, bFindInChild?true:false);
	}
	return S_OK;
}

STDMETHODIMP CWndPage::AttachObjEvent(IDispatch* HTMLWindow,IDispatch* SourceObj, BSTR bstrEventName,BSTR AliasName)
{
	CString strEventName = OLE2T(bstrEventName);
	if (strEventName == _T(""))
		return S_FALSE;
	CString strAliasName = OLE2T(AliasName);
	if (theApp.m_mapEventType.size() == 0)
	{
		theApp.InitEventDic();
	}

	auto it = theApp.m_mapEventType.find(strEventName);
	if (it == theApp.m_mapEventType.end())
		return S_FALSE;

	CComQIPtr<IHTMLWindow2> pWindow2(HTMLWindow);
	if (pWindow2 == NULL)
		return S_FALSE;

	for (auto it0 : m_mapJSExtender)
	{
		CComPtr<IHTMLWindow2> pWindow;
		it0.first->get_parentWindow(&pWindow);
		if (pWindow.p == pWindow2.p)
		{
			if (theApp.m_pCloudAddinCLRProxy)
			{
				theApp.m_pEventProxy = new CComObject < CEventProxy >;
				CComBSTR bstrName(L"");
				bstrName = theApp.m_pCloudAddinCLRProxy->AttachObjEvent((IEventProxy*)theApp.m_pEventProxy, SourceObj, it->second, HTMLWindow);
				if (strAliasName == _T(""))
					strAliasName = OLE2T(bstrName);
				return it0.second->SinkObject(strAliasName, theApp.m_pEventProxy);
			}
			break;
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_Handle(long* pVal)
{
	if (m_hWnd)
		*pVal = (long)m_hWnd;
	return S_OK;
}

STDMETHODIMP CWndPage::GetCtrlInNode(BSTR NodeName, BSTR CtrlName, IDispatch** ppCtrl)
{
	CString strName = OLE2T(NodeName);
	if (strName == _T(""))
		return S_OK;
	auto it2 = m_mapNode.find(strName);
	if (it2 == m_mapNode.end())
		return S_OK;

	CWndNode* pNode = it2->second;
	if (pNode)
		pNode->GetCtrlByName(CtrlName, true, ppCtrl);

	return S_OK;
}

STDMETHODIMP CWndPage::AttachNodeSubCtrlEvent(IDispatch* HtmlWndDisp, VARIANT Node, VARIANT CtrlName, BSTR EventName, BSTR AliasName)
{
	CComQIPtr<IHTMLWindow2> pWindow2(HtmlWndDisp);
	if (pWindow2 == nullptr)
		return S_FALSE;

	if (theApp.m_mapEventType.size() == 0)
	{
		theApp.InitEventDic();
	}
	CString strEventName = OLE2T(EventName);
	if (strEventName == _T(""))
		return S_FALSE;
	
	auto it1 = theApp.m_mapEventType.find(strEventName);
	if (it1 == theApp.m_mapEventType.end())
		return S_FALSE;

	CString strName = _T("");
	CWndNode* pNode = nullptr;
	if (Node.vt == VT_BSTR)
	{
		strName = OLE2T(Node.bstrVal);
		if (strName == _T(""))
			return S_OK;
	}
	else if (Node.vt == VT_DISPATCH)
	{
		CComQIPtr<IWndNode> _pNode(Node.pdispVal);
		if (_pNode)
		{
			pNode = (CWndNode*)_pNode.p;
			if (pNode->m_nViewType != CLRCtrl)
				return S_FALSE;
			strName = pNode->m_strWebObjName;
		}
		else
			return S_OK;
	}
	else
		return S_OK;
	if (strName == _T("") && pNode==nullptr)
		return S_OK;
	CString strCtrlName = _T("");
	CComPtr<IDispatch> pDisp;
	if (CtrlName.vt == VT_BSTR)
		strCtrlName = OLE2T(CtrlName.bstrVal);
	else if (CtrlName.vt == VT_DISPATCH)
		pDisp.Attach(CtrlName.pdispVal);
	if (strCtrlName == _T("") && pDisp==nullptr)
		return S_FALSE;

	if (pNode == nullptr)
	{
		auto it2 = m_mapNode.find(strName);
		if (it2 == m_mapNode.end())
			return S_OK;
		pNode = it2->second;
	}
	if (pNode)
	{
		if (theApp.m_pCloudAddinCLRProxy)
		{
			if (pDisp)
			{
				CComBSTR bstrName(L"");
				bstrName = theApp.m_pCloudAddinCLRProxy->GetCtrlName(pDisp);
				strCtrlName = OLE2T(bstrName);
			}
		}
		else
			return S_OK;

		if (pDisp==nullptr)
			pNode->GetCtrlByName(strCtrlName.AllocSysString(), true, &pDisp);

		if (pDisp)
		{
			CString strAliasName = OLE2T(AliasName);
			if (strAliasName == _T(""))
			{
				strAliasName = strName;
				strAliasName += _T("_");
				strAliasName += strCtrlName;
				strAliasName += _T("_");
				strAliasName += strEventName;
			}

			CJSExtender* _pJSExtender = nullptr;

			for (auto it0 : m_mapJSExtender)
			{
				CComPtr<IHTMLWindow2> pWindow;
				it0.first->get_parentWindow(&pWindow);
				if (pWindow.p == pWindow2.p)
				{
					_pJSExtender = it0.second;
					break;
				}
			}
			if (_pJSExtender == nullptr)
				return S_FALSE;
			theApp.m_pEventProxy = new CComObject < CEventProxy >;
			CComBSTR bstrName(L"");
			bstrName = theApp.m_pCloudAddinCLRProxy->AttachObjEvent((IEventProxy*)theApp.m_pEventProxy, pDisp, it1->second, HtmlWndDisp);
			return _pJSExtender->SinkObject(strAliasName, theApp.m_pEventProxy);
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_xtml(BSTR strKey, BSTR* pVal)
{
	map<CString, CString>::iterator it = m_mapXtml.find(OLE2T(strKey));
	if (it != m_mapXtml.end())
		*pVal = it->second.AllocSysString();

	return S_OK;
}

STDMETHODIMP CWndPage::put_xtml(BSTR strKey, BSTR newVal)
{
	CString strkey = OLE2T(strKey);
	CString strVal = OLE2T(newVal);
	if (strkey == _T("") || strVal == _T(""))
		return S_OK;
	auto it = m_mapXtml.find(strkey);
	if (it != m_mapXtml.end())
		m_mapXtml.erase(it);

	m_mapXtml[strkey] = strVal;
	return S_OK;
}

STDMETHODIMP CWndPage::Extend(IDispatch* Parent, BSTR CtrlName, BSTR FrameName, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode)
{
	if (theApp.m_pCloudAddinCLRProxy)
	{
		IDispatch* pDisp = nullptr;
		pDisp =theApp.m_pCloudAddinCLRProxy->GetCtrlByName(Parent, CtrlName, true);
		if (pDisp)
		{
			HWND h = 0;
			h = theApp.m_pCloudAddinCLRProxy->IsCtrlCanNavigate(pDisp);
			if (h)
			{
				CString strFrameName = OLE2T(FrameName);
				CString strKey = OLE2T(bstrKey);
				if (strFrameName == _T(""))
					FrameName = CtrlName;
				if (strKey == _T(""))
					bstrKey = CComBSTR(L"Default");
				auto it = m_mapFrame.find((HWND)h);
				if (it == m_mapFrame.end())
				{
					IWndFrame* pFrame = nullptr;
					CreateFrame(CComVariant(0), CComVariant((long)h), FrameName, &pFrame);
					return pFrame->Extend(bstrKey, bstrXml, ppRetNode);
				}
				else
				{
					return it->second->Extend(bstrKey, bstrXml, ppRetNode);
				}
			}
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::ExtendCtrl(VARIANT MdiForm, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode)
{
	HWND h = 0;
	bool bMDI = false;
	if (MdiForm.vt == VT_DISPATCH)
	{
		if (theApp.m_pCloudAddinCLRProxy)
		{
			h = theApp.m_pCloudAddinCLRProxy->GetMDIClientHandle(MdiForm.pdispVal);
			if (h)
				bMDI = true;
			else
			{
				h = theApp.m_pCloudAddinCLRProxy->IsCtrlCanNavigate(MdiForm.pdispVal);
				if (h)
				{
					CComBSTR bstrName(L"");
					bstrName = theApp.m_pCloudAddinCLRProxy->GetCtrlName(MdiForm.pdispVal);
					CString strKey = OLE2T(bstrKey);
					if (strKey == _T(""))
						bstrKey = CComBSTR(L"Default");
					IWndFrame* pFrame = nullptr;
					map<HWND, CWndFrame*>::iterator it = m_mapFrame.find((HWND)h);
					if (it == m_mapFrame.end())
					{
						CreateFrame(CComVariant(0), CComVariant((long)h), bstrName, &pFrame);
						return pFrame->Extend(bstrKey, bstrXml, ppRetNode);
					}
					else
					{
						return it->second->Extend(bstrKey, bstrXml, ppRetNode);
					}
				}
			}
		}
	}
	else if (MdiForm.vt == VT_I4 || MdiForm.vt == VT_I8)
	{
		HWND hWnd = NULL;
		if (MdiForm.vt == VT_I4)
		{
			if (MdiForm.lVal == 0)
			{
				hWnd = ::FindWindowEx(m_hWnd, NULL, _T("MDIClient"), NULL);
				if (hWnd)
				{
					bMDI = true;
				}
				else
				{
					hWnd = ::GetWindow(m_hWnd, GW_CHILD);
				}
			}
			else
			{
				hWnd = (HWND)MdiForm.lVal;
				if (::IsWindow(hWnd) == false)
					hWnd = ::GetWindow(m_hWnd, GW_CHILD);
			}
		}
		if (MdiForm.vt == VT_I8)
		{
			if (MdiForm.llVal == 0)
			{
				hWnd = ::FindWindowEx(m_hWnd, NULL, _T("MDIClient"), NULL);
				if (hWnd)
				{
					bMDI = true;
				}
				else
				{
					hWnd = ::GetWindow(m_hWnd, GW_CHILD);
				}
			}
			else
			{
				hWnd = (HWND)MdiForm.llVal;
				if (::IsWindow(hWnd) == false)
					hWnd = ::GetWindow(m_hWnd, GW_CHILD);

			}
		}
		h = hWnd;
	}
	if (h)
	{
		CString strKey = OLE2T(bstrKey);
		if (strKey == _T(""))
			bstrKey = CComBSTR(L"Default");
		IWndFrame* pFrame = nullptr;
		if (bMDI)
		{
			HRESULT hr = CreateFrame(CComVariant(0), CComVariant((long)h), CComBSTR(L"TangramMdiFrame"), &pFrame);
			if (pFrame)
			{
				return pFrame->Extend(bstrKey, bstrXml, ppRetNode);
			}
		}
		else
		{
			CString strKey = OLE2T(bstrKey);
			if (strKey == _T(""))
				bstrKey = CComBSTR(L"Default");

			IWndFrame* pFrame = nullptr;
			auto it = m_mapFrame.find((HWND)h);
			if (it == m_mapFrame.end())
			{
				TCHAR szBuffer[MAX_PATH];
				::GetWindowText((HWND)h, szBuffer, MAX_PATH);
				CString strText = szBuffer;
				if (strText == _T(""))
				{
					CString s = _T("");
					s.Format(_T("Frame%p"), h);
					strText = s;
				}
				CreateFrame(CComVariant(0), CComVariant((long)h), CComBSTR(strText), &pFrame);
				return pFrame->Extend(bstrKey, bstrXml, ppRetNode);
			}
			else
			{
				return it->second->Extend(bstrKey, bstrXml, ppRetNode);
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_EnableSinkCLRCtrlCreatedEvent(VARIANT_BOOL* pVal)
{
	*pVal = m_bEnableSinkCLRCtrlCreatedEvent;
	return S_OK;
}

STDMETHODIMP CWndPage::put_EnableSinkCLRCtrlCreatedEvent(VARIANT_BOOL newVal)
{
	m_bEnableSinkCLRCtrlCreatedEvent = newVal;
	return S_OK;
}

STDMETHODIMP CWndPage::OnChanged(DISPID dispID)
{
	if (DISPID_READYSTATE == dispID)
	{
		VARIANT vResult = { 0 };
		EXCEPINFO excepInfo;
		UINT uArgErr;
		DISPPARAMS dp = { NULL, NULL, 0, 0 };
		if (SUCCEEDED(m_pHtmlDocument2->Invoke(DISPID_READYSTATE, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dp, &vResult, &excepInfo, &uArgErr)))
		{
			assert(VT_I4 == V_VT(&vResult));
			m_lReadyState = (READYSTATE)V_I4(&vResult);
			switch (m_lReadyState)
			{
			case READYSTATE_COMPLETE:
				PostThreadMessage(m_nThreadId, WM_TANGRAM_WEBNODEDOCCOMPLETE, (WPARAM)0, (LPARAM)0);
				break;
			default:
				break;
			}

			VariantClear(&vResult);
		}
	}

	return NOERROR;
}

STDMETHODIMP CWndPage::OnRequestEdit(DISPID dispID)
{
	return NOERROR;
}

HRESULT CWndPage::Init(CString szURL)
{
	HRESULT hr;
	LPCONNECTIONPOINTCONTAINER pCPC = nullptr;
	CComPtr<IDispatch> pScriptDisp;

	if (szURL == _T(""))
	{
		m_nScheme = INTERNET_SCHEME_HTTP;
	}
	else
	{
		URL_COMPONENTS urlComponents;
		DWORD dwFlags = 0;
		m_nScheme = INTERNET_SCHEME_UNKNOWN;
		ZeroMemory((PVOID)&urlComponents, sizeof(URL_COMPONENTS));
		urlComponents.dwStructSize = sizeof(URL_COMPONENTS);

		if (szURL)
		{
			if (InternetCrackUrl(szURL, 0, dwFlags, &urlComponents))
			{
				m_nScheme = urlComponents.nScheme;
			}
		}
	}

	if (FAILED(hr = CoCreateInstance(CLSID_HTMLDocument, nullptr, CLSCTX_INPROC_SERVER, IID_IHTMLDocument2, (LPVOID*)&m_pHtmlDocument2)))
	{
		goto Error;
	}

	if (SUCCEEDED(hr = m_pHtmlDocument2->get_parentWindow(&m_pHTMLWindow2)))
	{

	}

	if (FAILED(hr = m_pHtmlDocument2->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&pCPC)))
	{
		goto Error;
	}

	if (FAILED(hr = pCPC->FindConnectionPoint(IID_IPropertyNotifySink, &m_pCP)))
	{
		goto Error;
	}

	m_hrConnected = m_pCP->Advise((LPUNKNOWN)(IPropertyNotifySink*)this, &m_dwCookie);

	if (m_pHtmlDocument2->get_Script(&pScriptDisp) == S_OK)
	{
		CComPtr<IDispatchEx> _pScriptEx;
		hr = pScriptDisp->QueryInterface<IDispatchEx>(&_pScriptEx);
		if (hr == S_OK)
		{
			ConnectJSEngine(_pScriptEx);
			CJSExtender::AddObject(_T("Tangram"), (IWndPage*)theApp.m_pTangram);
		}
	}

	m_mapJSExtender[m_pHtmlDocument2] = this;
	CJSExtender::AddObject(_T("WndPage"), (IWndPage*)this);
	SinkObject(_T("WndPage_On"), (IWndPage*)this);

	if (m_strURL.CompareNoCase(_T("about:blank")) == 0)
		m_bIsBlank = true;

	switch (m_nScheme)
	{
	case INTERNET_SCHEME_HTTP:
	case INTERNET_SCHEME_FTP:
	case INTERNET_SCHEME_RES:
	case INTERNET_SCHEME_GOPHER:
	case INTERNET_SCHEME_HTTPS:
	case INTERNET_SCHEME_FILE:
		// load URL using IPersistMoniker
		hr = LoadURLFromMoniker();
		break;
	case INTERNET_SCHEME_NEWS:
	case INTERNET_SCHEME_MAILTO:
	case INTERNET_SCHEME_SOCKS:
		// we don't handle these
		return E_FAIL;
		break;
	default:
		if (m_bIsBlank)
		{
			hr = LoadURLFromMoniker();
			break;
		}
		hr = LoadURLFromFile();
		break;
	}

	if (SUCCEEDED(hr) || E_PENDING == hr)
	{
		MSG msg;
		while (GetMessage(&msg, NULL, 0, 0))
		{
			if (WM_TANGRAM_WEBNODEDOCCOMPLETE == msg.message && NULL == msg.hwnd)
			{
				theApp.SetHook(::GetCurrentThreadId());
				::PostMessage(m_hWnd, WM_TANGRAM_WEBNODEDOCCOMPLETE, 1, 0);
				return 1;
			}
			else
			{
				DispatchMessage(&msg);
			}
		}
	}
	return S_OK;
Error:
	if (pCPC)
		pCPC->Release();

	return hr;
}

HRESULT CWndPage::LoadURLFromMoniker()
{
	HRESULT hr;
	LPMONIKER pMk			= nullptr;
	LPBINDCTX pBCtx			= nullptr;
	LPPERSISTMONIKER pPMk	= nullptr;

	if (FAILED(hr = CreateURLMonikerEx(nullptr, CComBSTR(m_strURL), &pMk, URL_MK_UNIFORM)))
	{
		return hr;
	}

	if (FAILED(hr = CreateBindCtx(0, &pBCtx)))
	{
		goto Error;
	}

	if (SUCCEEDED(hr = m_pHtmlDocument2->QueryInterface(IID_IPersistMoniker, (LPVOID*)&pPMk)))
	{
		hr = pPMk->Load(false, pMk, pBCtx, STGM_READ);
	}

Error:
	if (pMk) pMk->Release();
	if (pBCtx) pBCtx->Release();
	if (pPMk) pPMk->Release();
	return hr;
}

HRESULT CWndPage::LoadURLFromFile()
{
	HRESULT hr;
	LPPERSISTFILE  pPF;
	if (SUCCEEDED(hr = m_pHtmlDocument2->QueryInterface(IID_IPersistFile, (LPVOID*)&pPF)))
	{
		hr = pPF->Load(CComBSTR(m_strURL), 0);
		pPF->Release();
	}

	return hr;
}


HRESULT CWndPage::UnLoad()
{
	HRESULT hr = NOERROR;

	if (m_pCP)
	{
		if (m_dwCookie)
		{
			hr = m_pCP->Unadvise(m_dwCookie);
			m_dwCookie = 0;
		}

		m_pCP->Release();
		m_pCP = nullptr;
	}

	if (m_pHtmlDocument2)
	{
		CCommonFunction::ClearObject<IUnknown>(m_pHtmlDocument2);
		m_pHtmlDocument2 = nullptr;
	}

	if (m_pHTMLWindow2)
	{
		DWORD dw = m_pHTMLWindow2->Release();
		while(dw>2)
			dw = m_pHTMLWindow2->Release();
		//CCommonFunction::ClearObject<IUnknown>(m_pHTMLWindow2);
		m_pHTMLWindow2 = nullptr;
	}

	return NOERROR;
}


