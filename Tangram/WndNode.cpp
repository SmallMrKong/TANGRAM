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
*
********************************************************************************/

// WndNode.cpp : Implementation of CWndNode

#include "stdafx.h"
#include "CloudAddinApp.h"
#include "CloudAddinCore.h"
#include "NodeWnd.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "TangramHtmlTreeWnd.h"
#include "TangramCoreEvents.h"
#include "SplitterWnd.h"

// CWndNodeWBEvent
CWndNodeWBEvent::CWndNodeWBEvent(CWndNode* pNode)
{
	m_pWndNode			= pNode;
	m_pHTMLDocument2	= nullptr;
	DispEventAdvise(m_pWndNode->m_pDisp);
}

void CWndNodeWBEvent::OnNavigateComplete2(IDispatch* pDisp, VARIANT* URL)
{
	if (m_pWndNode->m_pDisp == pDisp)
	{
		ConnectJSEngine(NULL);
		CWndPage* pPage = m_pWndNode->m_pFrame->m_pPage;
		if (m_pHTMLDocument2)
		{
			map<IHTMLDocument2*, CJSExtender*>::iterator it = pPage->m_mapJSExtender.find(m_pHTMLDocument2);
			if (it != pPage->m_mapJSExtender.end())
			{
				pPage->m_mapJSExtender.erase(it);
			}
		}
		CComQIPtr<IWebBrowser2> pWeb2(pDisp);
		CComPtr<IDispatch> pDispDoc;
		pWeb2->get_Document(&pDispDoc);
		CComQIPtr<IHTMLDocument2> pHtmlDocument2(pDispDoc);
		if (pHtmlDocument2&&pPage)
		{
			m_pHTMLDocument2 = pHtmlDocument2.Detach();
			pPage->m_mapJSExtender[m_pHTMLDocument2] = this;
			CComPtr<IDispatch> pScriptDisp;
			CComPtr<IDispatchEx> _pScriptEx;
			if (m_pHTMLDocument2->get_Script(&pScriptDisp) == S_OK&&pScriptDisp->QueryInterface<IDispatchEx>(&_pScriptEx) == S_OK)
			{
				ConnectJSEngine(_pScriptEx);
				AddObject(_T("Tangram"), theApp.m_pTangram);
				AddObject(_T("WndPage"), (IWndPage*)pPage);
				SinkObject(_T("WndPage_On"), (IWndPage*)pPage);

				for (auto it2 : pPage->m_mapExternalObj)
				{
					AddObject(it2.first, (IDispatch*)it2.second);
					SinkObject(it2.first + _T("_On"), (IWndPage*)it2.second);
				}
			}
		}
	}
}

void CWndNodeWBEvent::OnDocumentComplete(IDispatch *pDisp, VARIANT *URL)
{
	if (m_pWndNode->m_pDisp == pDisp)
	{
		CWndPage* pPage = m_pWndNode->m_pPage;
		if (pPage)
		{
			bool bState = false;

			for (auto it: pPage->m_mapFrame)
			{
				for (auto it2 : it.second->m_mapNode)
				{
					if (it2.second->m_bCreated == false)
					{
						::PostMessage(pPage->m_hWnd, WM_TANGRAM_WEBNODEDOCCOMPLETE, 0, 0);
						return;
					}
				}
			}
			m_pWndNode->m_bNodeDocComplete = true;
			CComPtr<IHTMLWindow2> pWinDisp;
			m_pHTMLDocument2->get_parentWindow(&pWinDisp);
			pPage->Fire_PageLoaded(pWinDisp.p, CComBSTR(m_pWndNode->m_strURL));
		}
	}
}

CWndNode::CWndNode()
{
#ifdef _DEBUG
	theApp.m_nTangramObj++;
#endif

	m_nID						= 0;
	m_nRow						= 0;
	m_nCol						= 0;
	m_nRows						= 1;
	m_nCols						= 1;
	m_nViewType					= BlankView;

	m_bTopObj					= false;
	m_bWebInit					= false;
	m_bCreated					= false;
	m_bNodeDocComplete			= false;
	m_varTag.vt					= VT_EMPTY;
	m_strKey					= _T("");
	m_strURL					= _T("");
	m_strWebObjName				= _T("");
	m_strExtenderID				= _T("");
	m_hHostWnd					= NULL;
	m_hChildHostWnd				= NULL;
	m_pOfficeObj				= nullptr;
	m_pDisp						= nullptr;
	m_pFrame					= nullptr;
	m_pPage						= nullptr;
	m_pHostFrame				= nullptr;
	m_pRootObj					= nullptr;
	m_pHostWnd					= nullptr;
	m_pExtender					= nullptr;
	m_pParentObj				= nullptr;
	m_pObjClsInfo				= nullptr;
	m_pVisibleXMLObj			= nullptr;
	m_pHostClientView			= nullptr; 
	m_pCLREventConnector		= nullptr;
	m_pTangramNodeWBEvent		= nullptr;
	m_pChildNodeCollection		= nullptr;
	m_pCurrentExNode			= nullptr;
	m_pHostParse				= nullptr;
	m_pDocTreeCtrl				= nullptr;
	m_pTangramParse				= nullptr;
	m_pDocXmlParseNode			= nullptr;
	theApp.m_pWndNode			= this;
}

CWndNode::~CWndNode()
{
	if (theApp.m_pWndNode == this)
		theApp.m_pWndNode = nullptr;
	if (m_pPage)
	{
		auto it = m_pPage->m_mapNode.find(m_strWebObjName);
		if (it != m_pPage->m_mapNode.end())
			m_pPage->m_mapNode.erase(it);
	}

	if (m_strKey!=_T(""))
	{
		auto it = m_pFrame->m_mapNode.find(m_strKey);
		if (it != m_pFrame->m_mapNode.end())
		{
			m_pFrame->m_mapNode.erase(it);
			if (m_pFrame->m_mapNode.size() == 0)
			{
				if (::IsWindow(m_pFrame->m_hWnd))
				{
					m_pFrame->UnsubclassWindow(true);
					m_pFrame->m_hWnd = NULL;
				}
				delete m_pFrame;
			}
		}
	}

	if (m_pCLREventConnector)
	{
		HRESULT hr = m_pCLREventConnector->DispEventUnadvise((IWndNode*)this);
		delete m_pCLREventConnector;
		m_pCLREventConnector = nullptr;
	}

#ifdef _DEBUG
	theApp.m_nTangramObj--;
#endif
	HRESULT hr = S_OK;
	DWORD dw = 0;

	if (m_pExtender)
	{
		dw = m_pExtender->Release();
		m_pExtender = nullptr;
	}

	if (m_pTangramParse)
	{
		delete m_pTangramParse;
	}

	if (m_nViewType != TreeView&&m_nViewType != Splitter&&m_pDisp)
		CCommonFunction::ClearObject<IUnknown>(m_pDisp);

	m_vChildNodes.clear();

	if (m_pChildNodeCollection != nullptr)
	{
		dw = m_pChildNodeCollection->Release();
		m_pChildNodeCollection = nullptr;
	}

	ATLTRACE(_T("delete CWndNode:%x\n"), this);
}

BOOL CWndNode::PreTranslateMessage(MSG* pMsg)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	if(m_pHostWnd&&pMsg->message != WM_MOUSEMOVE)
	{
		m_pHostWnd->PreTranslateMessage(pMsg);

		for(auto it : m_vChildNodes)
		{
			it->PreTranslateMessage(pMsg);
		}
	}
	return true;
}

STDMETHODIMP CWndNode::LoadXML(int nType, BSTR bstrXML)
{
	if (m_strID.CompareNoCase(_T("TreeView")) == 0)
	{
		CTangramHtmlTreeWnd* _pXHtmlTree = ((CNodeWnd*)m_pHostWnd)->m_pXHtmlTree;
		if (nType)
		{
			_pXHtmlTree->LoadXmlFromFile(OLE2T(bstrXML), CTangramHtmlTreeWnd::ConvertToUnicode);
		}
		else
		{
			_pXHtmlTree->LoadXmlFromString(OLE2T(bstrXML), CTangramHtmlTreeWnd::ConvertToUnicode);
		}
	}

	return S_OK;
}

STDMETHODIMP CWndNode::UpdateDesignerData(BSTR bstrXML)
{
	if (m_pRootObj->m_pDocTreeCtrl)
	{
		CString strXml = OLE2T(bstrXML);
		m_pRootObj->m_pDocTreeCtrl->UpdateData(strXml);
	}

	return S_OK;
}

STDMETHODIMP CWndNode::ActiveTabPage(IWndNode* _pNode)
{
	theApp.m_pWndNode = this;
	if (::IsWindow(m_pHostWnd->m_hWnd))
		::SetFocus(m_pHostWnd->m_hWnd);
	theApp.HostPosChanged(this);
	return S_OK;
}

STDMETHODIMP CWndNode::Extend(BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode)
{
	if (m_pPage&&(m_nViewType==ActiveX||m_nViewType==CLRCtrl|| m_nViewType == BlankView))
	{
		if (m_pHostFrame == nullptr)
		{
			CString strName = m_strWebObjName;
			strName += _T("_Frame");

			if(m_nViewType == BlankView)
				m_hHostWnd = ::CreateWindowEx(NULL, L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, m_pHostWnd->m_hWnd, NULL, AfxGetInstanceHandle(), NULL);
			else
				m_hHostWnd = ::GetWindow(m_pHostWnd->m_hWnd,GW_CHILD);
			m_pPage->CreateFrame(CComVariant(0),CComVariant((long)m_hHostWnd),strName.AllocSysString(),&m_pHostFrame);
			((CWndFrame*)m_pHostFrame)->m_pHostNode = this;
		}
		if (m_pHostFrame&&::IsWindow(m_hHostWnd))
		{
			HRESULT hr = m_pHostFrame->Extend(bstrKey, bstrXml, ppRetNode);
			CWndNode* pRootNode = (CWndNode*)*ppRetNode;
			pRootNode->m_pHostFrame = m_pHostFrame;
			RECT rect;
			::GetWindowRect(pRootNode->m_pHostWnd->m_hWnd, &rect);
			HWND hWnd = ::GetParent(pRootNode->m_pHostWnd->m_hWnd);
			CWndNode* pNode = (CWndNode*)::SendMessage(hWnd, WM_TANGRAMMSG, 0, 0);
			if (pNode!= nullptr &&(pNode->m_nViewType == ViewType::TabbedWnd))
			{
				pNode->m_pHostWnd->ScreenToClient(&rect);
				((CWndFrame*)m_pHostFrame)->m_x = rect.left;
				((CWndFrame*)m_pHostFrame)->m_y = rect.top;
			}
			return hr;
		}
	}
	return S_OK;
}

STDMETHODIMP CWndNode::ExtendEx(int nRow, int nCol, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode)
{
	if (m_pPage&&m_nViewType == Splitter)
	{
		IWndNode* pNode = nullptr;
		GetNode(nRow, nCol, &pNode);
		if (pNode == nullptr)
			return S_OK;
		CWndNode* pWindowNode = (CWndNode*)pNode;
		if (pWindowNode->m_pHostFrame == nullptr)
		{
			CString strName = pWindowNode->m_strWebObjName;
			strName += _T("_Frame");

			::SetWindowLong(pWindowNode->m_pHostWnd->m_hWnd, GWL_ID, 1);
			m_pPage->CreateFrame(CComVariant(0), CComVariant((long)pWindowNode->m_pHostWnd->m_hWnd), strName.AllocSysString(), &pWindowNode->m_pHostFrame);
			((CWndFrame*)pWindowNode->m_pHostFrame)->m_pHostNode = this;
		}
		if (pWindowNode->m_pHostFrame)
		{
			if (pWindowNode->m_pCurrentExNode)
			{
				::SetWindowLong(pWindowNode->m_pCurrentExNode->m_pHostWnd->m_hWnd, GWL_ID, 1);
			}
			HRESULT hr = pWindowNode->m_pHostFrame->Extend(bstrKey, bstrXml, ppRetNode);
			if (hr != S_OK)
			{
				if (pWindowNode->m_pCurrentExNode)
					::SetWindowLong(pWindowNode->m_pCurrentExNode->m_pHostWnd->m_hWnd, GWL_ID, AFX_IDW_PANE_FIRST + nRow * 16 + nCol);
				else
					::SetWindowLong(pWindowNode->m_pHostWnd->m_hWnd, GWL_ID, AFX_IDW_PANE_FIRST + nRow * 16 + nCol);
				return S_OK;
			}
			CWndNode* pRootNode = (CWndNode*)*ppRetNode;
			pWindowNode->m_pCurrentExNode = pRootNode;
			::SetWindowLongPtr(pRootNode->m_pHostWnd->m_hWnd, GWLP_ID, AFX_IDW_PANE_FIRST + nRow * 16 + nCol);
			return hr;
		}
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Tag(VARIANT* pVar)
{
	*pVar = m_varTag;

	if (m_varTag.vt == VT_DISPATCH && m_varTag.pdispVal)
		m_varTag.pdispVal->AddRef();
	return S_OK;
}

STDMETHODIMP CWndNode::put_Tag(VARIANT var)
{
	m_varTag = var;
	return S_OK;
}

STDMETHODIMP CWndNode::get_XObject(VARIANT* pVar)
{
	pVar->vt = VT_EMPTY;
	if (m_pDisp)
	{
		pVar->vt = VT_DISPATCH;
		pVar->pdispVal = m_pDisp;
		pVar->pdispVal->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_AxPlugIn(BSTR bstrPlugInName, IDispatch** pVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString strObjName = OLE2T(bstrPlugInName);
	strObjName.Trim();
	strObjName.MakeLower();
	IDispatch* pDisp = nullptr;
	if(m_pRootObj->m_PlugInDispDictionary.Lookup(LPCTSTR(strObjName),(void*&)pDisp))
	{
		* pVal = pDisp;
		(* pVal)->AddRef();
	}
	else
		*pVal = nullptr;
	return S_OK;
} 

STDMETHODIMP CWndNode::get_Name(BSTR* pVal)
{
	*pVal = m_strName.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndNode::put_Name(BSTR bstrNewName)
{
	CString strName = OLE2T(bstrNewName);
	strName.Trim();
	strName.Replace(_T(","),_T(""));
	if (m_pHostParse != NULL)
	{
		CString _strName = _T(",");
		_strName += theApp.GetNames(this);
		CString _strName2 = _T(",");
		_strName2 += strName;
		_strName2 += _T(",");
		int nPos = _strName.Find(_strName2);
		if (nPos == -1)
		{
			m_pHostParse->put_attr(L"name", strName);
			m_strName = strName;
		}
		else
		{
			::MessageBox(NULL,_T("Modify name failed!"),_T("Tangram"),MB_OK);
		}
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Attribute(BSTR bstrKey, BSTR* pVal)
{
	if (m_pHostParse != nullptr)
	{
		CString strVal = m_pHostParse->attr(OLE2T(bstrKey), _T(""));
		*pVal = strVal.AllocSysString();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_Attribute(BSTR bstrKey, BSTR bstrVal)
{
	if (m_pHostParse != nullptr)
	{
		CString strID = OLE2T(bstrKey);
		CString strVal = OLE2T(bstrVal);
		m_pHostParse->put_attr(strID, strVal);
		if (strVal.CompareNoCase(_T("hostview")) == 0)
		{
			m_strID = _T("hostview");
			m_pHostWnd->EnableWindow(false);
			m_pHostClientView = (CNodeWnd*)m_pHostWnd;
			CWndNode* pTopNode = m_pRootObj;
			while(pTopNode != pTopNode->m_pRootObj)
			{
				pTopNode->m_pFrame->m_pBindingNode = this;
				pTopNode->m_pHostClientView = m_pHostClientView;
				pTopNode = pTopNode->m_pRootObj;
			}
			if (m_pFrame->m_pBindingNode&&m_pFrame->m_pBindingNode != this)
			{
				m_pFrame->m_pBindingNode->m_strID = _T("");
				m_pFrame->m_pBindingNode->put_Attribute(_T("id"),_T(""));
				m_pFrame->m_pBindingNode->m_pHostWnd->Invalidate();
			}
			m_pFrame->m_pBindingNode = this;
			m_pRootObj->m_pFrame->m_pBindingNode = this;
			m_pRootObj->m_pHostClientView = m_pHostClientView;

			if (m_pParentObj&&m_pParentObj->m_nViewType == Splitter)
				m_pHostWnd->ModifyStyleEx(WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE, 0);
			theApp.HostPosChanged(this);
		}
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Caption(BSTR* pVal)
{
	*pVal = m_strCaption.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndNode::put_Caption(BSTR bstrCaption)
{
	CString str(bstrCaption);

	m_strCaption = str;

	if (m_pParentObj != nullptr && m_pParentObj->m_pHostWnd != nullptr)
	{
		m_pParentObj->m_pHostWnd->SendMessage(WM_TGM_SET_CAPTION,m_nCol,(LPARAM)str.GetBuffer());
	}

	if (m_pHostParse != nullptr)
	{
		m_pHostParse->put_attr(L"caption",str);
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Handle(long* pVal)
{
	if(m_pHostWnd)	
		*pVal = (long)m_pHostWnd->m_hWnd;
	return S_OK;
}

STDMETHODIMP CWndNode::get_OuterXml(BSTR* pVal)
{
	*pVal = m_pRootObj->m_pTangramParse->xml().AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndNode::get_Key(BSTR* pVal)
{
	*pVal = m_pRootObj->m_strKey.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndNode::get_XML(BSTR* pVal)
{
	*pVal = m_pHostParse->xml().AllocSysString();
	return S_OK;
}

HWND CWndNode::CreateView(HWND hParentWnd, CString strTag)
{
	BOOL bWebCtrl = false;
	CString strURL = _T("");
	CString strID = strTag;
	auto hWnd = ::CreateWindowEx(NULL,L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE| WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, hParentWnd, NULL, AfxGetInstanceHandle(), NULL);
	CComBSTR bstr2;
	get_Attribute(CComBSTR("Name"), &bstr2);
	CString strName = OLE2T(bstr2);
	switch (m_nViewType)
	{
	case ActiveX:
		{
			strID.MakeLower();
			auto nPos = strID.Find(_T("http:"));
			if (nPos == -1)
				nPos = strID.Find(_T("https:"));
			if (m_pFrame)
			{
				CComBSTR bstr;
				get_Attribute(CComBSTR("InitInfo"), &bstr);
				CString str = _T("");
				str += bstr;
				if (str != _T("") && m_pPage)
				{
					LRESULT hr = ::SendMessage(m_pPage->m_hWnd, WM_GETNODEINFO, (WPARAM)OLE2T(bstr), (LPARAM)hParentWnd);
					if (hr)
					{
						CString strInfo = _T("");
						CWindow m_wnd;
						m_wnd.Attach(hParentWnd);
						m_wnd.GetWindowText(strInfo);
						strID += strInfo;
						m_wnd.Detach();
					}
				}
			}

			if (strID.Find(_T("http://")) != -1 || strID.Find(_T("https://")) != -1)
			{
				bWebCtrl = true;
				strURL = strID;
				strID = _T("shell.explorer.2");
			}

			if (strID.Find(_T("res://")) != -1 || ::PathFileExists(strID))
			{
				bWebCtrl = true;
				strURL = strID;
				strID = _T("shell.explorer.2");
			}

			if (strID.CompareNoCase(_T("about:blank")) ==0 )
			{
				bWebCtrl = true;
				strURL = strID;
				strID = _T("shell.explorer.2");
			}

			if (m_pDisp == nullptr)
			{
				CComPtr<IDispatch> pDisp;
				HRESULT hr = pDisp.CoCreateInstance(CComBSTR(strID));
				if (hr == S_OK)
				{
					m_pDisp = pDisp.Detach();
				}
			}
		}
		break;
	case CLRCtrl:
		{
			if (theApp.m_pCloudAddinCLRProxy)
				m_pDisp = theApp.m_pCloudAddinCLRProxy->TangramCreateObject(strTag.AllocSysString(), (long)hParentWnd, this);
		}
		break;
	}
	if (m_pDisp)
	{
		CAxWindow m_Wnd;
		m_Wnd.Attach(hWnd);
		CComPtr<IUnknown> pUnk;
		m_Wnd.AttachControl(m_pDisp, &pUnk);
		if (m_nViewType == ActiveX)
		{
			CComQIPtr<IWebBrowser2> pWebDisp(m_pDisp);
			if (pWebDisp)
			{
				bWebCtrl = true;
				m_strURL = strURL;
				if (m_strURL == _T(""))
					m_strURL = strID;
				if (m_pFrame)
				{
					if (m_pPage)
						m_pPage->m_nWebViewCount++;

					m_pTangramNodeWBEvent = new CWndNodeWBEvent(this);
				}
				CComPtr<IAxWinAmbientDispatch> spHost;
				LRESULT hr = m_Wnd.QueryHost(&spHost);
				if (SUCCEEDED(hr))
				{
					CComBSTR bstr;
					get_Attribute(CComBSTR("scrollbar"), &bstr);
					CString str = OLE2T(bstr);
					if (str == _T("1"))
						spHost->put_DocHostFlags(DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_THEME);//DOCHOSTUIFLAG_DIALOG|
					else
						spHost->put_DocHostFlags(/*DOCHOSTUIFLAG_DIALOG|*/DOCHOSTUIFLAG_SCROLL_NO | DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_THEME);

					if (m_strURL != _T(""))
					{
						pWebDisp->Navigate2(&CComVariant(m_strURL), &CComVariant(navNoReadFromCache | navNoWriteToCache), nullptr, nullptr, nullptr);
						m_bWebInit = true;
					}
				}
			}
		}

		CComQIPtr<IOleInPlaceActiveObject> pIOleInPlaceActiveObject(m_pDisp);
		if (pIOleInPlaceActiveObject)
			((CNodeWnd*)m_pHostWnd)->m_pOleInPlaceActiveObject = pIOleInPlaceActiveObject.Detach();
		m_Wnd.Detach();
		m_pRootObj->m_mapLayoutNodes[m_strName] = this;
	}

	return hWnd;
}

STDMETHODIMP CWndNode::get_ChildNodes(IWndNodeCollection** pNodeColletion)
{
	if (m_pChildNodeCollection == nullptr)
	{
		CComObject<CWndNodeCollection>::CreateInstance(&m_pChildNodeCollection);
		m_pChildNodeCollection->AddRef();		
		m_pChildNodeCollection->m_pNodes = &m_vChildNodes;
	}
	return m_pChildNodeCollection->QueryInterface(IID_IWndNodeCollection,(void**)pNodeColletion);
}

int CWndNode::_getNodes(CWndNode* pNode, CString& strName,CWndNode**ppRetNode, CWndNodeCollection* pNodes)
{
	int iCount = 0;
	if (pNode->m_strName.CompareNoCase(strName) == 0)
	{
		if (pNodes != nullptr)
			pNodes->m_pNodes->push_back(pNode);

		if (ppRetNode != nullptr && (*ppRetNode) == nullptr)
			*ppRetNode = pNode;
		return 1;
	}

	for(auto it : pNode->m_vChildNodes)
	{
		iCount += _getNodes(it,strName,ppRetNode, pNodes);
	}
	return iCount;
}

STDMETHODIMP CWndNode::Show()
{
	CWndNode* pChild = this;
	CWndNode* pParent = pChild->m_pParentObj;

	while(pParent != nullptr)
	{
		pParent->m_pHostWnd->SendMessage(WM_ACTIVETABPAGE,(WPARAM)pChild->m_nCol,(LPARAM)1);

		pChild = pParent;
		pParent = pChild->m_pParentObj;
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_RootNode(IWndNode** ppNode)
{
	if (m_pRootObj != nullptr)
		*ppNode = m_pRootObj;
	return S_OK;
}

STDMETHODIMP CWndNode::get_ParentNode(IWndNode** ppNode)
{
	*ppNode = nullptr;
	if (m_pParentObj != nullptr)
		*ppNode = m_pParentObj;

	return S_OK;
}

STDMETHODIMP CWndNode::get_NodeType(WndNodeType* nType)
{
	*nType = (WndNodeType)m_nViewType;
	return S_OK;
}

void CWndNode::_get_Objects(CWndNode* pNode, UINT32& nType, CWndNodeCollection* pNodeColletion)
{
	if (pNode->m_nViewType & nType)
	{
		pNodeColletion->m_pNodes->push_back(pNode);
	}

	CWndNode* pChildNode = nullptr;
	for(auto it : pNode->m_vChildNodes)
	{
		pChildNode = it;
		_get_Objects(pChildNode,nType,pNodeColletion);
	}
}

STDMETHODIMP CWndNode::get_Objects(long nType, IWndNodeCollection** ppNodeColletion)
{
	CComObject<CWndNodeCollection>* pNodes = nullptr;
	CComObject<CWndNodeCollection>::CreateInstance(&pNodes);

	pNodes->AddRef();	

	UINT32 u = nType;
	_get_Objects(this,u,pNodes);
	HRESULT hr = pNodes->QueryInterface(IID_IWndNodeCollection,(void**)ppNodeColletion);

	pNodes->Release();

	return hr;
}

STDMETHODIMP CWndNode::get_Rows(long* nRows)
{
	*nRows = m_nRows;
	return S_OK;
}

STDMETHODIMP CWndNode::get_Cols(long* nCols)
{
	*nCols = m_nCols;
	return S_OK;
}

STDMETHODIMP CWndNode::get_Row(long* nRow)
{
	*nRow = m_nRow;
	return S_OK;
}

STDMETHODIMP CWndNode::get_Col(long* nCol)
{
	*nCol = m_nCol;
	return S_OK;
}

STDMETHODIMP CWndNode::GetNode(long nRow, long nCol,IWndNode** ppTangramNode)
{
	CWndNode* pRet = nullptr;
	auto bFound = false;

	*ppTangramNode = nullptr;
	if (nRow < 0 || nCol < 0 || nRow >= m_nRows || nCol >= m_nCols) return E_INVALIDARG;

	for(auto it : m_vChildNodes)
	{
		pRet = it;
		if (pRet->m_nCol == nCol && pRet->m_nRow == nRow)
		{
			bFound = true;
			break;
		}
	}

	HRESULT hr = S_OK;
	if (bFound)
	{
		hr = pRet->QueryInterface(IID_IWndNode,(void**)ppTangramNode);
	}
	return hr;
}

STDMETHODIMP CWndNode::GetNodes(BSTR bstrName,IWndNode** ppNode,IWndNodeCollection** ppNodes,long* pCount)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString strName(bstrName);

	CWndNode* pRetNode = nullptr;

	if (ppNode != nullptr)
		*ppNode = nullptr;

	CComObject<CWndNodeCollection>* pNodes = nullptr;
	if (ppNodes != nullptr)
	{
		*ppNodes = nullptr;
		CComObject<CWndNodeCollection>::CreateInstance(&pNodes);
		pNodes->AddRef();
	}

	int iCount = _getNodes(this,strName,&pRetNode,pNodes);

	*pCount = iCount;

	if ((iCount > 0) && (ppNode != nullptr))
		pRetNode->QueryInterface(IID_IWndNode,(void**)ppNode);	

	if (ppNodes != nullptr)
		pNodes->QueryInterface(IID_IWndNodeCollection,(void**)ppNodes);

	if (pNodes != nullptr)
		pNodes->Release();

	return S_OK;
}

BOOL CWndNode::AddChildNode(CWndNode* pNode)
{
	m_vChildNodes.push_back(pNode);
	pNode->m_pParentObj = this;
	pNode->m_pFrame = m_pFrame;
	pNode->m_pRootObj = m_pRootObj;
	return true;
}

BOOL CWndNode::RemoveChildNode(CWndNode* pNode)
{
	auto it = find(m_vChildNodes.begin(), m_vChildNodes.end(), pNode);
	if (it != m_vChildNodes.end())
	{
		m_vChildNodes.erase(it);
		return true;
	}
	return false;
}

STDMETHODIMP CWndNode::get_Frame(IWndFrame** pVal)
{
	if(m_pFrame)
		*pVal = m_pFrame;

	return S_OK;
}

STDMETHODIMP CWndNode::Refresh(void)
{
	if(m_pDisp)
	{
		CComQIPtr<IWebBrowser2> pWebCtrl(m_pDisp);
		if(pWebCtrl)
			pWebCtrl->Refresh();
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_Height(LONG* pVal)
{
	RECT rc;
	::GetClientRect(m_pHostWnd->m_hWnd,&rc);
	*pVal = rc.bottom;
	return S_OK;
}

STDMETHODIMP CWndNode::get_Width(LONG* pVal)
{
	RECT rc;
	::GetClientRect(m_pHostWnd->m_hWnd, &rc);
	*pVal = rc.right;

	return S_OK;
}

STDMETHODIMP CWndNode::get_Extender(IDispatch** pVal)
{
	if(m_pExtender)
	{
		* pVal = m_pExtender;
		(* pVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_Extender(IDispatch* newVal)
{
	if(m_pExtender)
		m_pExtender->Release();
	m_pExtender = newVal;
	m_pExtender->AddRef();

	return S_OK;
}

STDMETHODIMP CWndNode::get_WndPage(IWndPage** pVal)
{
	*pVal = (IWndPage*)m_pRootObj->m_pFrame->m_pPage;
	return S_OK;
}

STDMETHODIMP CWndNode::get_HTMLWindow(IDispatch** pVal)
{
	if (m_pTangramNodeWBEvent&&m_pTangramNodeWBEvent->m_pHTMLDocument2)
	{
		CComPtr<IHTMLWindow2> pHTMLWindow2;
		m_pTangramNodeWBEvent->m_pHTMLDocument2->get_parentWindow(&pHTMLWindow2);
		*pVal = pHTMLWindow2.Detach();
		(*pVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_HtmlDocument(IDispatch** pVal)
{
	if (m_pTangramNodeWBEvent&&m_pTangramNodeWBEvent->m_pHTMLDocument2)
	{
		*pVal = m_pTangramNodeWBEvent->m_pHTMLDocument2;
		(*pVal)->AddRef();
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_NameAtWindowPage(BSTR* pVal)
{
	*pVal = m_strWebObjName.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndNode::GetCtrlByName(BSTR bstrName, VARIANT_BOOL bFindInChild, IDispatch** ppRetDisp)
{
	if (theApp.m_pCloudAddinCLRProxy&&m_pDisp)
		*ppRetDisp=theApp.m_pCloudAddinCLRProxy->GetCtrlByName(m_pDisp, bstrName, bFindInChild?true:false);

	return S_OK;
}

CWndNodeCollection::CWndNodeCollection()
{
	m_pNodes = &m_vNodes;
	theApp._addObject(this, this);
}

CWndNodeCollection::~CWndNodeCollection()
{
	theApp._removeObject(this);
	m_vNodes.clear();
}

STDMETHODIMP CWndNodeCollection::get_NodeCount(long* pCount)
{
	*pCount = (int)m_pNodes->size();
	return S_OK;
}

STDMETHODIMP CWndNodeCollection::get_Item(long iIndex, IWndNode **ppNode)
{
	if (iIndex < 0 || iIndex >= (int)m_pNodes->size()) return E_INVALIDARG;

	CWndNode* pNode = m_pNodes->operator [](iIndex);

	return pNode->QueryInterface(IID_IWndNode, (void**)ppNode);
}

STDMETHODIMP CWndNodeCollection::get__NewEnum(IUnknown** ppVal)
{
	*ppVal = nullptr;

	struct _CopyVariantFromIUnkown
	{
		static HRESULT copy(VARIANT* p1, CWndNode* const* p2)
		{
			CWndNode* pNode = *p2;
			p1->vt = VT_UNKNOWN;
			return pNode->QueryInterface(IID_IUnknown, (void**)&(p1->punkVal));
		}

		static void init(VARIANT* p)	{ VariantInit(p); }
		static void destroy(VARIANT* p)	{ VariantClear(p); }
	};

	typedef CComEnumOnSTL<IEnumVARIANT, &IID_IEnumVARIANT, VARIANT, _CopyVariantFromIUnkown, CTangramNodeVector>
		CComEnumVariantOnVector;

	CComObject<CComEnumVariantOnVector> *pe = 0;
	HRESULT hr = CComObject<CComEnumVariantOnVector>::CreateInstance(&pe);

	if (SUCCEEDED(hr))
	{
		hr = pe->AddRef();
		hr = pe->Init(GetUnknown(), *m_pNodes);

		if (SUCCEEDED(hr))
			hr = pe->QueryInterface(ppVal);

		hr = pe->Release();
	}

	return hr;
}

STDMETHODIMP CWndNode::get_DocXml(BSTR* pVal)
{
	CString strXml = m_pRootObj->m_pTangramParse->xml();
	*pVal = strXml.AllocSysString();
	strXml.ReleaseBuffer();

	return S_OK;
}

STDMETHODIMP CWndNode::get_rgbMiddle(OLE_COLOR* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = OLE_COLOR(pSplitter->rgbMiddle);
	}
	else
	{
		*pVal = OLE_COLOR(RGB(240,240,240));
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_rgbMiddle(OLE_COLOR newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		OleTranslateColor(newVal, NULL, &pSplitter->rgbMiddle);
		BYTE Red = GetRValue(pSplitter->rgbMiddle); 
		BYTE Green = GetGValue(pSplitter->rgbMiddle); 
		BYTE Blue = GetBValue(pSplitter->rgbMiddle); 
		CString strRGB = _T("");
		strRGB.Format(_T("RGB(%d,%d,%d)"), Red, Green, Blue);
		put_Attribute(CComBSTR(L"middlecolor"), strRGB.AllocSysString());
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_rgbLeftTop(OLE_COLOR* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = OLE_COLOR(pSplitter->rgbLeftTop);
	}
	else
	{
		*pVal = OLE_COLOR(RGB(240, 240, 240));
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_rgbLeftTop(OLE_COLOR newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		OleTranslateColor(newVal, NULL, &pSplitter->rgbLeftTop);
		CString strRGB = _T("");
		strRGB.Format(_T("RGB(%d,%d,%d)"), GetRValue(pSplitter->rgbLeftTop), GetGValue(pSplitter->rgbLeftTop), GetBValue(pSplitter->rgbLeftTop));
		put_Attribute(CComBSTR(L"lefttopcolor"), strRGB.AllocSysString());
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_rgbRightBottom(OLE_COLOR* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = OLE_COLOR(pSplitter->rgbRightBottom);
	}
	else
	{
		*pVal = OLE_COLOR(RGB(240, 240, 240));
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_rgbRightBottom(OLE_COLOR newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		OleTranslateColor(newVal, NULL, &pSplitter->rgbRightBottom);
		BYTE Red = GetRValue(pSplitter->rgbRightBottom);
		BYTE Green = GetGValue(pSplitter->rgbRightBottom);
		BYTE Blue = GetBValue(pSplitter->rgbRightBottom);
		CString strRGB = _T("");
		strRGB.Format(_T("RGB(%d,%d,%d)"), Red, Green, Blue);
		put_Attribute(CComBSTR(L"rightbottomcolor"), strRGB.AllocSysString());
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Hmin(int* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = pSplitter->m_Hmin;
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_Hmin(int newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		pSplitter->m_Hmin = min(pSplitter->m_Hmax, newVal);
		CString strVal = _T("");
		strVal.Format(_T("%d"), pSplitter->m_Hmin);
		put_Attribute(CComBSTR(L"hmin"), strVal.AllocSysString());
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_Hmax(int* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = pSplitter->m_Hmax;
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_Hmax(int newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		pSplitter->m_Hmax = max(pSplitter->m_Hmin,newVal);
		CString strVal = _T("");
		strVal.Format(_T("%d"), pSplitter->m_Hmax);
		put_Attribute(CComBSTR(L"hmax"), strVal.AllocSysString());
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_Vmin(int* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = pSplitter->m_Vmin;
	}

	return S_OK;
}

STDMETHODIMP CWndNode::put_Vmin(int newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		pSplitter->m_Vmin = min(pSplitter->m_Vmax, newVal);
		CString strVal = _T("");
		strVal.Format(_T("%d"), pSplitter->m_Vmin);
		put_Attribute(CComBSTR(L"vmin"), strVal.AllocSysString());
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_Vmax(int* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = pSplitter->m_Vmax;
	}

	return S_OK;
}

STDMETHODIMP CWndNode::put_Vmax(int newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		pSplitter->m_Vmax = max(pSplitter->m_Vmin, newVal);
		CString strVal = _T("");
		strVal.Format(_T("%d"), pSplitter->m_Vmax);
		put_Attribute(CComBSTR(L"vmax"), strVal.AllocSysString());
	}

	return S_OK;
}

