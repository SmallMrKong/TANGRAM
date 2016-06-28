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

// TangramView.cpp : implementation file
//

#include "stdafx.h"
#include "CloudAddinApp.h"
#include "NodeWnd.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "TangramHtmlTreeWnd.h"
#include "OfficePlus\ExcelPlus\ExcelAddin.h"
#include "OfficePlus\ExcelPlus\ExcelPlusWnd.h"

// CNodeWnd

IMPLEMENT_DYNCREATE(CNodeWnd, CWnd)

CNodeWnd::CNodeWnd()
{
	m_bBKWnd					= false;
	m_bCreateExternal			= false;
	m_bEraseBkgnd				= true;
	m_pWndNode					= nullptr;
	m_pXHtmlTree				= nullptr;
	m_pParentNode				= nullptr;
	m_pOleInPlaceActiveObject	= nullptr;
}

CNodeWnd::~CNodeWnd()
{
}

BEGIN_MESSAGE_MAP(CNodeWnd, CWnd)
	ON_WM_DESTROY()
	ON_WM_ERASEBKGND()
	ON_WM_MOUSEACTIVATE()
	ON_MESSAGE(WM_TANGRAMGETNODE, OnGetTangramObj)
	ON_MESSAGE(WM_TGM_SETACTIVEPAGE,OnActiveTangramObj)
	ON_MESSAGE(WM_TANGRAMMSG,OnTangramMsg)
	ON_MESSAGE(WM_TABCHANGE,OnTabChange)
	ON_WM_WINDOWPOSCHANGED()
	ON_WM_SIZE()
END_MESSAGE_MAP()

// CNodeWnd diagnostics

#ifdef _DEBUG
void CNodeWnd::AssertValid() const
{
	///CView::AssertValid();
}

#ifndef _WIN32_WCE
void CNodeWnd::Dump(CDumpContext& dc) const
{
	CWnd::Dump(dc);
}
#endif
#endif //_DEBUG

 //CNodeWnd message handlers
BOOL CNodeWnd::Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext)
{
	m_pWndNode = theApp.m_pWndNode;
	m_pWndNode -> m_nID = nID;
	m_pWndNode -> m_pHostWnd = this;

	if(m_pWndNode->m_strID.CompareNoCase(_T("HostView"))==0)
	{
		CWndFrame* pWnd = m_pWndNode->m_pFrame;
		pWnd->m_pBindingNode = m_pWndNode;

		m_pWndNode->m_pFrame->m_pWorkNode->m_pHostClientView = this;

		HWND hWnd = CreateWindow(L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, pParentWnd->m_hWnd, (HMENU)nID, AfxGetInstanceHandle(), NULL);
		BOOL bRet = SubclassWindow(hWnd);
		if(m_pWndNode->m_pParentObj&&m_pWndNode->m_pParentObj->m_nViewType==Splitter)
			ModifyStyleEx(WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE, 0);

		BSTR bstrVal = ::SysAllocString(L"");
		m_pWndNode->get_Attribute(CComBSTR(_T("mdiclient")), &bstrVal);
		CString strVal = OLE2T(bstrVal);
		if (strVal.CompareNoCase(_T("true")) == 0)
		{
			BOOL bEnableBK = false;
			LRESULT lRes = ::SendMessage(m_pWndNode->m_pFrame->m_hWnd,WM_TANGRAMMSG,0,1992);//Query Support BK WebPage
			bEnableBK = (lRes==0);
			if(bEnableBK)
			{
				BSTR bstrURL = ::SysAllocString(L"");
				m_pWndNode->get_Attribute(CComBSTR(_T("url")), &bstrURL);
				CString strURL = OLE2T(bstrURL);
				::SysFreeString(bstrURL);
				if (strURL != _T(""))
				{
					if (strURL.Find(_T("//")) == -1 && ::PathFileExists(strURL) == false)
					{
						CString strPath = theApp.m_strAppPath + _T("TangramWebPage\\") + strURL;
						if (::PathFileExists(strPath))
							strURL = strPath;
					}

					m_bBKWnd = true;
					if (pWnd->m_pBKWnd == nullptr)
					{
						pWnd->m_pBKWnd = new CBKWnd();
						pWnd->m_pBKWnd->m_pWndNode = m_pWndNode;
						CComPtr<IDispatch> pDisp2;
						HRESULT hr = pDisp2.CoCreateInstance(CComBSTR(_T("shell.explorer.2")));
						CAxWindow m_wnd;
						m_wnd.Attach(pWnd->m_hWnd);
						CComPtr<IUnknown> pUnk;
						m_wnd.AttachControl(pDisp2, &pUnk);
						theApp._addObject(this, pUnk.Detach());
						pWnd->m_pBKWnd->SubclassWindow(::FindWindowEx(pWnd->m_hWnd, NULL, _T("Shell Embedding"), NULL));
						theApp.m_pMDIClientBKWnd = pWnd->m_pBKWnd;
						m_pWndNode->m_pDisp = pDisp2.Detach();
						CComQIPtr<IWebBrowser2> pWebDisp(m_pWndNode->m_pDisp);
						if (pWebDisp)
						{
							CComQIPtr<IOleInPlaceActiveObject> pIOleInPlaceActiveObject(m_pWndNode->m_pDisp);
							if (pIOleInPlaceActiveObject)
								m_pOleInPlaceActiveObject = pIOleInPlaceActiveObject.Detach();
							CComPtr<IAxWinAmbientDispatch> spHost;
							hr = m_wnd.QueryHost(&spHost);
							if (SUCCEEDED(hr))
							{
								hr = spHost->put_DocHostFlags(DOCHOSTUIFLAG_DIALOG  | DOCHOSTUIFLAG_DISABLE_HELP_MENU| DOCHOSTUIFLAG_SCROLL_NO | DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_THEME);
								spHost->put_AllowContextMenu(false);
							}
							pWebDisp->Navigate2(&CComVariant(strURL), &CComVariant(navNoReadFromCache | navNoWriteToCache), nullptr, nullptr, nullptr);
							CWndPage* pPage = m_pWndNode->m_pFrame->m_pPage;
							if (pPage)
								pPage->m_nWebViewCount++;

							m_pWndNode->m_pTangramNodeWBEvent = new CWndNodeWBEvent(m_pWndNode);
						}
						m_wnd.Detach();
					}
				}
			}
		}
		::SysFreeString(bstrVal);
		return bRet ;
	}

	return theApp.Create(m_pWndNode,lpszClassName, lpszWindowName, dwStyle, rect, pParentWnd, nID, pContext);
}

int CNodeWnd::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message)
{
	CWndFrame* pFrame = m_pWndNode->m_pRootObj->m_pFrame;
	BOOL b = pFrame->m_bDesignerState;
	if (theApp.m_pDocDOMTree&&theApp.m_pCloudAddinCLRProxy)
	{
		if (b)
			theApp.m_pCloudAddinCLRProxy->SelectNode(m_pWndNode);
	}
	if (m_pWndNode&&m_pWndNode->m_pPage)
		m_pWndNode->m_pPage->Fire_NodeMouseActivate(m_pWndNode);

	if((m_pWndNode->m_nViewType==TabbedWnd || m_pWndNode->m_nViewType==Splitter))
	{
		if (theApp.m_pWndFrame != m_pWndNode->m_pFrame)
			::SetFocus(m_hWnd);
		theApp.m_pWndNode = m_pWndNode;
		theApp.m_bWinFormActived = false;
		return MA_ACTIVATE;
	}
		
	if(!m_bCreateExternal)
		Invalidate(true);
	else
	{
		if((m_pWndNode->m_nViewType!=TabbedWnd && m_pWndNode->m_nViewType!=Splitter))
		{
			if (theApp.m_pWndFrame != m_pWndNode->m_pFrame||theApp.m_pWndNode != m_pWndNode)
				::SetFocus(m_hWnd);
		}
	}
	theApp.m_pWndNode = m_pWndNode;
	theApp.m_bWinFormActived = false;
	if(m_pWndNode->m_pParentObj)
	{
		if (m_pWndNode->m_pParentObj->m_nViewType & TabbedWnd)
			m_pWndNode->m_pParentObj->m_pVisibleXMLObj = m_pWndNode;
	}
	theApp.m_pWndFrame = m_pWndNode->m_pFrame;

	if ((m_pWndNode->m_nViewType == ActiveX || m_pWndNode->m_nViewType == CLRCtrl))
	{
		return MA_ACTIVATE;
	}

	if (b)
	{
		theApp.m_pHostCore->CreateCommonDesignerToolBar();
		if (m_pWndNode->m_strID.CompareNoCase(_T("hostview")) && (m_bCreateExternal == false && m_pWndNode->m_pDisp == NULL))
		{
			if (::IsWindow(theApp.m_pHostCore->m_hHostWnd)==false)
			{
				theApp.m_pHostCore->m_hHostWnd = ::CreateWindowEx(NULL, L"Tangram Window Class", _T("Tangram Designer Helper Window"), WS_OVERLAPPED | WS_CAPTION, 0, 0, 0, 0, NULL, NULL, theApp.m_hInstance, NULL);
				theApp.m_mapValInfo[_T("hostwindow")] = CComVariant((long)theApp.m_pHostCore->m_hHostWnd);
			}
			if (theApp.m_pDesignWindowNode&&theApp.m_pDesignWindowNode!=m_pWndNode)
			{
				CNodeWnd* pWnd = ((CNodeWnd*)theApp.m_pDesignWindowNode->m_pHostWnd);
				if (pWnd&&::IsWindow(pWnd->m_hWnd))
				{
					pWnd->Invalidate(true);
				}
			}
			theApp.m_pDesignWindowNode = m_pWndNode;
			Invalidate(true);
		}

		HWND hFrame = m_pWndNode->m_pFrame->m_hWnd;
		HWND hWnd = (HWND)::SendMessage(hFrame, WM_TANGRAMGETDESIGNWND,0,0);

		if (m_bCreateExternal == false)
		{
			IWndNode* pTopDesignNode = (IWndNode*)::SendMessage(pFrame->m_hWnd, WM_TANGRAMISDOCUMENT, 0, 0);
			if (pTopDesignNode)
			{
				theApp.m_pHostDesignUINode = (CWndNode*)pTopDesignNode;
				theApp.m_pDesignWindowNode = m_pWndNode;
			}
		}
		if (theApp.m_pDesigningFrame != pFrame)
		{
			theApp.m_pHostDesignUINode = m_pWndNode->m_pRootObj;
			theApp.m_pDesigningFrame = pFrame;
			pFrame->UpdateDesignerTreeInfo();
		}
	}

	if(m_bCreateExternal==false)
		return MA_ACTIVATEANDEAT;
	else
		return CWnd::OnMouseActivate(pDesktopWnd, nHitTest, message);
}

BOOL CNodeWnd::OnEraseBkgnd(CDC* pDC)
{
	BOOL bInDesignState = m_pWndNode->m_pRootObj->m_pFrame->m_bDesignerState;
	if (m_pWndNode->m_strID.CompareNoCase(_T("HostView")) == 0)
	{
		HWND hWnd = m_pWndNode->m_pFrame->m_hWnd;
		if (m_pWndNode->m_pRootObj->m_pFrame->m_bDesignerState)
		{
			RECT rt;
			CBitmap bit;
			GetClientRect(&rt);
			bit.LoadBitmap(IDB_BITMAP_MDI);
			CBrush br(&bit);
			pDC->FillRect(&rt, &br);
		}
		if (::IsWindow(hWnd) && !::IsWindowVisible(hWnd))
		{
			theApp.HostPosChanged(m_pWndNode);
			return true;
		}
	}

	if (m_pWndNode->m_strID.CompareNoCase(_T("hostview")) && (m_bCreateExternal == false && m_pWndNode->m_pDisp == NULL)&& m_bEraseBkgnd)
	{
		RECT rt;
		CBitmap bit;
		GetClientRect(&rt);
		if (bInDesignState&&theApp.m_pDesignWindowNode== m_pWndNode)
		{
			bit.LoadBitmap(IDB_BITMAP_DESIGNER);
			CBrush br(&bit);
			pDC->FillRect(&rt, &br);
		}
		else
		{
			bit.LoadBitmap(IDB_BITMAP_GRID);
			CBrush br(&bit);
			pDC->FillRect(&rt,&br);
			pDC->SetBkMode(TRANSPARENT);
			CComBSTR bstrCaption(L"");
			m_pWndNode->get_Attribute(CComBSTR(L"caption"), &bstrCaption);
			CString strInfo = _T("");
			strInfo = strInfo +
				_T("\n  ----Tangram Object Information----") +
				_T("\n  ") +
				_T("\n   Object Name:   ") + m_pWndNode->m_strName +
				_T("\n   Object Caption:") + OLE2T(bstrCaption) +
				_T("\n  ") +
				_T("\n  ");

			pDC->SetTextColor(RGB(255,255,255));
			pDC->DrawText(strInfo,&rt,DT_LEFT);
		}
	}
	return true;
}

BOOL CNodeWnd::PreTranslateMessage(MSG* pMsg)
{
	if (m_pXHtmlTree)
	{
		return m_pXHtmlTree->PreTranslateMessage(pMsg);
	}

	if(m_pOleInPlaceActiveObject)
	{
		LRESULT hr = m_pOleInPlaceActiveObject->TranslateAccelerator((LPMSG)pMsg);
		if(hr==S_OK)
			return true;
		else
		{
			if (m_pWndNode->m_nViewType == CLRCtrl)
			{
				VARIANT_BOOL bShiftKey = false;
				if (::GetKeyState(VK_SHIFT) < 0)  // shift pressed
					bShiftKey = true;
				hr = theApp.m_pCloudAddinCLRProxy->ProcessCtrlMsg(::GetWindow(m_hWnd, GW_CHILD), bShiftKey?true:false);
				if(hr==S_OK)
					return true;
			}
			else
			{
				pMsg->hwnd = 0;
				return true;
			}
		}
	}
	if (IsDialogMessage(pMsg))
		return true;
	return CWnd::PreTranslateMessage(pMsg);
}

void CNodeWnd::OnDestroy()
{
	if (theApp.m_pDesignWindowNode == m_pWndNode)
	{
		if (theApp.m_pCloudAddinCLRProxy)
			theApp.m_pCloudAddinCLRProxy->SelectNode(NULL);
		theApp.m_pDesignWindowNode = NULL;
	}

	m_pWndNode->Fire_Destroy();
	CWnd::OnDestroy();
}

void CNodeWnd::PostNcDestroy()
{
	CWnd::PostNcDestroy();
	delete m_pWndNode;
	delete this;
}

LRESULT CNodeWnd::OnTabChange(WPARAM wParam, LPARAM lParam)
{
	m_pWndNode->m_nActivePage = (int)wParam;
	IWndNode* pNode = nullptr;
	m_pWndNode->GetNode(0, wParam, &pNode);
	if (pNode&&theApp.m_pDocDOMTree&&theApp.m_pCloudAddinCLRProxy)
	{
		if(pNode)
			theApp.m_pCloudAddinCLRProxy->SelectNode(pNode);
	}
	return CWnd::DefWindowProc(WM_TABCHANGE, wParam, lParam);
}

LRESULT CNodeWnd::OnTangramMsg(WPARAM wParam, LPARAM lParam)
{	
	if (wParam == 0 && lParam)//Create CLRCtrl Node
	{
		CString strCnnID = (LPCTSTR)lParam;
		if (strCnnID.Find(_T(","))!=-1&&theApp.m_pCloudAddinCLRProxy)
		{
			m_pWndNode->m_pDisp = theApp.m_pCloudAddinCLRProxy->TangramCreateObject(strCnnID.AllocSysString(), (long)::GetParent(m_hWnd), m_pWndNode);

			if (m_pWndNode->m_pDisp)
			{
				CString strName = theApp.m_pCloudAddinCLRProxy->m_strObjTypeName;
				CWndNode* pNode = m_pWndNode->m_pRootObj;
				CWndNode* pParent = m_pWndNode->m_pParentObj;
				if (pParent)
				{
					strName = pParent->m_strName + _T("_") + strName;
				}
				auto it = pNode->m_mapLayoutNodes.find(strName);
				if (it != pNode->m_mapLayoutNodes.end())
				{
					BOOL bGetNew = false;
					int nIndex = 0;
					while (bGetNew == false)
					{
						CString strNewName = _T("");
						strNewName.Format(_T("%s%d"), strName, nIndex);
						it = pNode->m_mapLayoutNodes.find(strNewName);
						if (it == pNode->m_mapLayoutNodes.end())
						{
							strName = strNewName;
							break;
						}
						nIndex++;
					}
				}
				m_pWndNode->put_Attribute(CComBSTR(L"name"), strName.AllocSysString());
				m_pWndNode->m_strName = strName;
				m_pWndNode->m_pRootObj->m_mapLayoutNodes[strName] = m_pWndNode;
				m_pWndNode->m_nViewType = CLRCtrl;
				CAxWindow m_Wnd;
				m_Wnd.Attach(m_hWnd);
				CComPtr<IUnknown> pUnk;
				m_Wnd.AttachControl(m_pWndNode->m_pDisp, &pUnk);
				CComQIPtr<IOleInPlaceActiveObject> pIOleInPlaceActiveObject(m_pWndNode->m_pDisp);
				if (pIOleInPlaceActiveObject)
					m_pOleInPlaceActiveObject = pIOleInPlaceActiveObject.Detach();
				m_Wnd.Detach();
			}
		}
		else
		{
			BOOL bWebCtrl = false;
			CString strURL = _T("");
			strCnnID.MakeLower();
			auto nPos = strCnnID.Find(_T("http:"));
			if (nPos == -1)
				nPos = strCnnID.Find(_T("https:"));
			if (m_pWndNode->m_pFrame)
			{
				CComBSTR bstr;
				m_pWndNode->get_Attribute(CComBSTR("InitInfo"), &bstr);
				CString str = _T("");
				str += bstr;
				if (str != _T("") && m_pWndNode->m_pPage)
				{
					LRESULT hr = ::SendMessage(m_pWndNode->m_pPage->m_hWnd, WM_GETNODEINFO, (WPARAM)OLE2T(bstr), (LPARAM)::GetParent(m_hWnd));
					if (hr)
					{
						CString strInfo = _T("");
						CWindow m_wnd;
						m_wnd.Attach(::GetParent(m_hWnd));
						m_wnd.GetWindowText(strInfo);
						strCnnID += strInfo;
						m_wnd.Detach();
					}
				}
			}

			if (strCnnID.Find(_T("http://")) != -1 || strCnnID.Find(_T("https://")) != -1)
			{
				bWebCtrl = true;
				strURL = strCnnID;
				strCnnID = _T("shell.explorer.2");
			}

			if (strCnnID.Find(_T("res://")) != -1 || ::PathFileExists(strCnnID))
			{
				bWebCtrl = true;
				strURL = strCnnID;
				strCnnID = _T("shell.explorer.2");
			}

			if (strCnnID.CompareNoCase(_T("about:blank")) == 0)
			{
				bWebCtrl = true;
				strURL = strCnnID;
				strCnnID = _T("shell.explorer.2");
			}

			if (m_pWndNode->m_pDisp == NULL)
			{
				CComPtr<IDispatch> pDisp;
				HRESULT hr = pDisp.CoCreateInstance(CComBSTR(strCnnID));
				if (hr == S_OK)
				{
					m_pWndNode->m_pDisp = pDisp.Detach();
				}
			}
			if (m_pWndNode->m_pDisp)
			{
				m_pWndNode->m_pRootObj->m_mapLayoutNodes[m_pWndNode->m_strName] = m_pWndNode;
				m_pWndNode->m_nViewType = ActiveX;
				CAxWindow m_Wnd;
				m_Wnd.Attach(m_hWnd);
				CComPtr<IUnknown> pUnk;
				m_Wnd.AttachControl(m_pWndNode->m_pDisp, &pUnk);
				CComQIPtr<IWebBrowser2> pWebDisp(m_pWndNode->m_pDisp);
				if (pWebDisp)
				{
					bWebCtrl = true;
					m_pWndNode->m_strURL = strURL;
					if (m_pWndNode->m_strURL == _T(""))
						m_pWndNode->m_strURL = strCnnID;
					if (m_pWndNode->m_pFrame)
					{
						if (m_pWndNode->m_pPage)
							m_pWndNode->m_pPage->m_nWebViewCount++;

						m_pWndNode->m_pTangramNodeWBEvent = new CWndNodeWBEvent(m_pWndNode);
					}
					CComPtr<IAxWinAmbientDispatch> spHost;
					LRESULT hr = m_Wnd.QueryHost(&spHost);
					if (SUCCEEDED(hr))
					{
						CComBSTR bstr;
						m_pWndNode->get_Attribute(CComBSTR("scrollbar"), &bstr);
						CString str = OLE2T(bstr);
						if (str == _T("1"))
							spHost->put_DocHostFlags(DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_THEME);//DOCHOSTUIFLAG_DIALOG|
						else
							spHost->put_DocHostFlags(/*DOCHOSTUIFLAG_DIALOG|*/DOCHOSTUIFLAG_SCROLL_NO | DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_THEME);

						if (m_pWndNode->m_strURL != _T(""))
						{
							pWebDisp->Navigate2(&CComVariant(m_pWndNode->m_strURL), &CComVariant(navNoReadFromCache | navNoWriteToCache), NULL, NULL, NULL);
							m_pWndNode->m_bWebInit = true;
						}
					}
				}
				CComQIPtr<IOleInPlaceActiveObject> pIOleInPlaceActiveObject(m_pWndNode->m_pDisp);
				if (pIOleInPlaceActiveObject)
					m_pOleInPlaceActiveObject = pIOleInPlaceActiveObject.Detach();
				m_Wnd.Detach();
			}
		}
		return 0;
	}
	CTangramXmlParse* pNewNXmlode = (CTangramXmlParse*)wParam;
	if (pNewNXmlode == nullptr)
	{
		return 0;
	}
	CString str = (LPCTSTR)lParam;
	CString strID = pNewNXmlode->attr(_T("id"), _T(""));
	IWndNode* _pNode = nullptr;
	CWndNode* pOldNode = (CWndNode*)m_pWndNode;
	if (m_pWndNode->m_hHostWnd == 0)
	{
		RECT rc;
		::GetClientRect(m_hWnd, &rc);
		m_pWndNode->m_hHostWnd = ::CreateWindowEx(NULL, L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, rc.right, rc.bottom, m_hWnd, NULL, AfxGetInstanceHandle(), NULL);
		m_pWndNode->m_hChildHostWnd = ::CreateWindowEx(NULL, L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, rc.right, rc.bottom, m_pWndNode->m_hHostWnd, NULL, AfxGetInstanceHandle(), NULL);
		IWndFrame* pFrame = nullptr;
		CString strName = pNewNXmlode->attr(_T("name"), _T(""))+_T("frame");
		m_pWndNode->m_pPage->CreateFrame(CComVariant(0), CComVariant((long)pOldNode->m_hChildHostWnd), strName.AllocSysString(), &pFrame);
		IWndNode* pNode = nullptr;
		pFrame->Extend(CComBSTR(L""),str.AllocSysString(),&pNode);
		m_bEraseBkgnd = false;

		((CWndNode*)pNode)->m_pRootObj = m_pWndNode->m_pRootObj;
		((CWndNode*)pNode)->m_pParentObj = m_pWndNode->m_pParentObj;
		m_pWndNode->m_pRootObj->m_mapLayoutNodes[((CWndNode*)pNode)->m_strName] = (CWndNode*)pNode;
		CString strXml = ((CWndNode*)pNode)->m_pHostParse->xml();
		CTangramXmlParse* pNew = ((CWndNode*)pNode)->m_pHostParse;
		CTangramXmlParse* pOld = pOldNode->m_pHostParse;
		CTangramXmlParse* pParent = m_pWndNode->m_pHostParse->m_pParentParse;
		CTangramXmlParse* pRet = nullptr;
		if (pParent)
		{
			pRet = pParent->ReplaceNode(pOld, pNew, _T(""));
			CString str = pRet->xml();
			int nCount = pRet->GetCount();
			((CWndNode*)pNode)->m_pHostParse = pRet;
			m_pWndNode->m_pHostParse = pRet;

			CWndNode* pChildNode = nullptr;
			for (auto it2 : ((CWndNode*)pNode)->m_vChildNodes)
			{
				pChildNode = it2;
				pChildNode->m_pRootObj = m_pWndNode->m_pRootObj;
				CString strName = pChildNode->m_strName;
				for (int i = 0; i < nCount; i++)
				{
					CTangramXmlParse* child = pRet->GetChild(i);
					CString _strName = child->attr(_T("name"),_T(""));
					if (_strName.CompareNoCase(strName) == 0)
					{
						pChildNode->m_pHostParse = child;
						break;
					}
				}
			}
			m_pWndNode->m_vChildNodes.push_back(((CWndNode*)pNode));
		}

		strXml = m_pWndNode->m_pRootObj->m_pTangramParse->xml();
		theApp.m_pHostDesignUINode = m_pWndNode->m_pRootObj;
		if (theApp.m_pHostDesignUINode)
		{
			CTangramHtmlTreeWnd* pTreeCtrl = theApp.m_pDocDOMTree;
			pTreeCtrl->DeleteItem(theApp.m_pDocDOMTree->m_hFirstRoot);

			if (pTreeCtrl->m_pHostXmlParse)
			{
				delete pTreeCtrl->m_pHostXmlParse;
			}
			pTreeCtrl->m_pHostXmlParse = new CTangramXmlParse();
			pTreeCtrl->m_pHostXmlParse->LoadXml(strXml);
			pTreeCtrl->m_hFirstRoot = pTreeCtrl->LoadXmlFromXmlParse(pTreeCtrl->m_pHostXmlParse);
			pTreeCtrl->ExpandAll();
		}
	}
	return -1;
}

LRESULT CNodeWnd::OnActiveTangramObj(WPARAM wParam, LPARAM lParam)
{
	if(m_pWndNode->m_nViewType==CLRCtrl)
		::SetWindowLong(m_hWnd,GWL_STYLE,WS_CHILD|WS_VISIBLE|WS_CLIPCHILDREN|WS_CLIPSIBLINGS);
	
	theApp.HostPosChanged(m_pWndNode);
	return -1;
}

LRESULT CNodeWnd::OnGetTangramObj(WPARAM wParam, LPARAM lParam)
{
	if (m_pWndNode)
		return (LRESULT)m_pWndNode;
	return (long)0;
}

LRESULT CNodeWnd::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	if(m_bCreateExternal)
	{
		switch(message)
		{
		case WM_COMMAND:
		{
			WNDPROC* lplpfn = GetSuperWndProcAddr();
			LRESULT res = CallWindowProc(*lplpfn, m_hWnd, message, wParam, lParam);
			HWND hWnd = (HWND)lParam;
			if (::IsWindow(hWnd)&&m_pWndNode)
			{
				::GetClassName(hWnd, theApp.m_szBuffer, MAX_PATH);
				m_pWndNode->Fire_ControlNotify(m_pWndNode, HIWORD(wParam), LOWORD(wParam), lParam, CComBSTR(theApp.m_szBuffer));
			}

			return res;
		}
		case WM_ACTIVATE:
		case WM_ERASEBKGND:
		case WM_SETFOCUS:
		case WM_CONTEXTMENU:
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
		case WM_LBUTTONUP:
		case WM_RBUTTONUP:
		case WM_LBUTTONDBLCLK:
			{
				WNDPROC* lplpfn = GetSuperWndProcAddr();
				return CallWindowProc(*lplpfn,m_hWnd,message, wParam, lParam);
			}
		case WM_MOUSEACTIVATE:
			{
				break;
			}
		case WM_SHOWWINDOW:
			{
				if (wParam&&m_pWndNode->m_strURL!=_T(""))
				{
					CComQIPtr<IWebBrowser2> pWebCtrl(m_pWndNode->m_pDisp);
					if(pWebCtrl)
					{
						if (m_pWndNode->m_bWebInit == false)
						{
							pWebCtrl->Navigate2(&CComVariant(m_pWndNode->m_strURL), &CComVariant(navNoReadFromCache | navNoWriteToCache), NULL, NULL, NULL);
							m_pWndNode->m_bWebInit = true;
						}
					}
				}
				break;
			}
		}
	}

	return CWnd::WindowProc(message, wParam, lParam);
}

CBKWnd::CBKWnd(void)
{
	m_pWndNode = nullptr; 
}

void CBKWnd::OnFinalMessage(HWND hWnd)
{
	CWindowImpl<CBKWnd, CWindow>::OnFinalMessage(hWnd);
	delete this;
}

LRESULT CBKWnd::OnMouseActivate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	if (m_pWndNode)
	{
		theApp.m_pWndNode = m_pWndNode;
		theApp.m_pWndFrame = m_pWndNode->m_pFrame;
		theApp.m_bWinFormActived = false;
	}
	return 1;
}

LRESULT CBKWnd::OnWindowPosChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	WINDOWPOS* lpwndpos = (WINDOWPOS*)lParam;
	lpwndpos->flags |= SWP_NOZORDER |SWP_NOACTIVATE | SWP_NOREDRAW;

	return DefWindowProc(uMsg, wParam, lParam);
}

void CNodeWnd::OnWindowPosChanged(WINDOWPOS* lpwndpos)
{

	CWnd::OnWindowPosChanged(lpwndpos);

	if (m_pWndNode&&m_pWndNode->m_nViewType == BlankView&&m_pWndNode->m_hHostWnd)
	{
		::SetWindowPos(m_pWndNode->m_hHostWnd, HWND_BOTTOM,0,0,lpwndpos->cx,lpwndpos->cy,SWP_NOACTIVATE|SWP_NOREDRAW);
		if(m_pWndNode->m_hChildHostWnd)
			::SetWindowPos(m_pWndNode->m_hChildHostWnd, HWND_BOTTOM,0,0,lpwndpos->cx,lpwndpos->cy,SWP_NOACTIVATE|SWP_NOREDRAW);
	}
	if (m_pWndNode->m_strID.CompareNoCase(_T("hostview")) && (m_bCreateExternal == false && m_pWndNode->m_pDisp == NULL))
	{
		return;
	}
	if (m_pWndNode->m_nViewType == TreeView)
	{
		lpwndpos->flags &= ~SWP_NOREDRAW;
		::SetWindowPos(m_pXHtmlTree->m_hWnd, NULL, 0, 0, lpwndpos->cx, lpwndpos->cy, SWP_NOACTIVATE);
	}
	if (this->m_bCreateExternal)
	{
		Invalidate();
	}
	if (m_pWndNode->m_pHostFrame)
	{
		if (m_pParentNode == NULL)
		{
			HWND hWnd = ::GetParent(m_pWndNode->m_pHostWnd->m_hWnd);
			CWndNode* pParentNode = (CWndNode*)::SendMessage(hWnd, WM_TANGRAMMSG, 0, 0);
			if (pParentNode&&pParentNode->m_nViewType == ViewType::TabbedWnd)
			{
				m_pParentNode = pParentNode;
			}

		}
		if (m_pParentNode)
		{
			lpwndpos->x -= ((CWndFrame*)m_pWndNode->m_pHostFrame)->m_x;
			lpwndpos->y -= ((CWndFrame*)m_pWndNode->m_pHostFrame)->m_y;
			::SetWindowPos(m_hWnd, HWND_BOTTOM,lpwndpos->x,lpwndpos->y,lpwndpos->cx,lpwndpos->cy,SWP_NOACTIVATE|SWP_NOREDRAW| SWP_NOSENDCHANGING);
		}
	}
}

void CNodeWnd::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);

	if (m_pWndNode->m_strID.CompareNoCase(_T("hostview")) && (m_bCreateExternal == false && m_pWndNode->m_pDisp == NULL))
	{
		return;
	}
	if (m_pWndNode->m_nViewType == TreeView&&m_pWndNode->m_pParentObj == NULL)
	{
		::SetWindowPos(m_pXHtmlTree->m_hWnd, NULL, 0, 0, cx, cy,/*SWP_NOREDRAW|*/SWP_NOACTIVATE);
	}
}
