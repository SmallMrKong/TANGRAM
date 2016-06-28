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

// TangramCore.cpp : Implementation of CTangram

#include "stdafx.h"
#include "CloudAddinCore.h"
#include "CloudAddinApp.h"
#include "DocTemplateDlg.h"
#include "EclipsePlus\EclipseAddin.h"

#include "OfficePlus\CloudAddin.h"
#include "OfficePlus\LyncPlus\lync.h"
#include "OfficePlus\ExcelPlus\Excel.h"
#include "OfficePlus\WordPlus\MSWord.h"
#include "OfficePlus\ProjectPlus\MSPrj.h"
#include "OfficePlus\OutLookPlus\MsOutl.h"
#include "OfficePlus\PowerpointPlus\msppt.h"
#include "CloudUtilities\CloudAddinDownLoad.h"
#include "VisualStudioPlus\CloudAddinVSAddin.h"

#include "RestObj.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "WebTaskObj.h"
#include "TangramHtmlTreeWnd.h"
//#include "include\cef_app.h"
class ATL_NO_VTABLE CTangramHelper : 
	public CComObjectRootBase,
	public CComCoClass<CTangramHelper>
{
public:
	static HRESULT WINAPI UpdateRegistry(BOOL bRegister)
	{
		return theApp.UpdateRegistryFromResource(IDR_TANGRAM, bRegister);
	}

	static HRESULT WINAPI CreateInstance(void* pv, REFIID riid, LPVOID* ppv)
	{
		if (theApp.m_pTangram)
		{
			DWORD dwID = ::GetCurrentThreadId();
			if (dwID != theApp.m_dwThreadID)
			{
				IStream* pStream = 0;
				HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_IDispatch, theApp.m_pTangram, &pStream);
				if (hr == S_OK)
				{
					IDispatch* pTarget = nullptr;
					hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_IDispatch, (LPVOID *)&pTarget);
					if (hr == S_OK&&pTarget)
					{
						*ppv = pTarget;
						return S_OK;
					}
				}
			}
			theApp.m_pTangram->QueryInterface(riid, ppv);
		}
		//*ppv = theApp.m_pTangram;
		return S_OK;
	}
};

TANGRAM_OBJECT_ENTRY_AUTO(CLSID_Tangram, CTangramHelper)

// CTangram

CTangram::CTangram()
{
	m_bCanClose			= false;
	m_nRef				= 4;
	m_hMainWnd			= NULL;
	m_hMainFrameWnd		= NULL;
	m_hHostWnd			= NULL;
	m_hChildHostWnd		= NULL;
	m_pPage				= nullptr;
	m_pFrame			= nullptr;
	m_pRootNodes		= nullptr;
	theApp.m_pHostCore	= this;
	TRACE(_T("------------------Create CTangram------------------------\n"));
	m_strDefaultTemplateXml = _T("");
}

void CTangram::Init()
{
	ITypeLib* pTypeLib = nullptr;
	ITypeInfo* pTypeInfo = nullptr;
	GetTypeInfo(0, 0, &pTypeInfo);
	if (pTypeInfo == nullptr)
		GetTI(0, &pTypeInfo);
	if (pTypeInfo)
	{
		pTypeInfo->GetContainingTypeLib(&pTypeLib, 0);
		pTypeLib->GetTypeInfoOfGuid(DIID__IEventProxy, &theApp.m_pEventTypeInfo);
		pTypeInfo->Release();
		pTypeLib->Release();
	}
}

CTangram::~CTangram()
{
	auto it = m_mapWebTask.begin();
	while (it != m_mapWebTask.end())
	{
		if (it->second->m_hHandle == 0)
			delete it->second;
		else
		{
			::WaitForSingleObject(it->second->m_hHandle, INFINITE);
			::PostThreadMessage(it->second->m_dwThreadID, WM_QUIT, 0, 0);
		}
		it = m_mapWebTask.begin();
	}

	if (::IsWindow(m_hHostWnd))
	{
		::DestroyWindow(m_hHostWnd);
	}

	for (auto it : m_mapObjDic)
	{
		if(it.first.CompareNoCase(_T("dte")))
			it.second->Release();
	}
	m_mapObjDic.erase(m_mapObjDic.begin(), m_mapObjDic.end());
	if (m_pRootNodes)
		CCommonFunction::ClearObject<CWndNodeCollection>(m_pRootNodes);
	TRACE(_T("------------------Release CTangram------------------------\n"));
	theApp.m_pTangram = nullptr;
	theApp.m_pHostCore = nullptr;
}

ULONG CTangram::InternalRelease()
{
	if (m_bCanClose == false)
		return 1;
	else if (::GetModuleHandle(_T("kso.dll")))
	{
		m_nRef--;
		return m_nRef;
	}
	return 1;
}

void CTangram::CreateCommonDesignerToolBar()
{
	if (theApp.m_pTangramPackageProxy)
	{
		theApp.m_pTangramPackageProxy->CreateTangramToolWnd();
		return;
	}
	if (::IsWindow(m_hHostWnd) == false)
	{
		m_hHostWnd = ::CreateWindowEx(WS_EX_PALETTEWINDOW, _T("Tangram Window Class"), _T("Tangram Designer"), WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, 0, 0, 300, 700, NULL, 0, theApp.m_hInstance, NULL);
		m_hChildHostWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), _T(""), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, m_hHostWnd, 0, theApp.m_hInstance, NULL);
	}

	if (m_hHostWnd)
	{
		auto it = theApp.m_mapValInfo.find(_T("tangramdesignertoolbar"));
		if (it != theApp.m_mapValInfo.end())
		{
			CreateWndPage((long)m_hHostWnd, &theApp.m_pDesignerWndPage);
			IWndFrame* pFrame = nullptr;
			theApp.m_pDesignerWndPage->CreateFrame(CComVariant(0), CComVariant((long)m_hChildHostWnd), CComBSTR(L"DesignerToolBar"), &pFrame);
			theApp.m_pDesignerFrame = (CWndFrame*)pFrame;
			IWndNode* pNode = nullptr;
			theApp.m_pDesignerFrame->Extend(CComBSTR(L""), it->second.bstrVal, &pNode);
		}
		HWND hwnd = ::GetActiveWindow();
		RECT rc;
		::GetWindowRect(hwnd, &rc);
		::SetWindowPos(m_hHostWnd, NULL, rc.left + 40, rc.top + 40, 300, 700, SWP_NOACTIVATE | SWP_NOREDRAW);
		::ShowWindow(m_hHostWnd, SW_SHOW);
		::UpdateWindow(m_hHostWnd);
	}
}

CWndNode* CTangram::ExtendEx(long hWnd, CString strXml, LONG l, LONG t, LONG r, LONG b, LONG l2, LONG t2, LONG r2, LONG b2)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	CTangramXmlParse* m_pParse = new CTangramXmlParse();
	bool bXml = m_pParse->LoadXml(strXml);
	if (bXml == false)
		bXml = m_pParse->LoadFile(strXml);

	if (bXml == false)
	{
		delete m_pParse;
		return nullptr;
	}

	BOOL bSizable = m_pParse->attrBool(_T("sizable"), false);
	CTangramXmlParse* pWndNode = m_pParse->GetChild(_T("window"));
	if (pWndNode == nullptr)
	{
		delete m_pParse;
		return nullptr;
	}

	CTangramXmlParse* pNode = pWndNode->GetChild(_T("node"));
	if (pNode == nullptr)
	{
		delete m_pParse;
		return nullptr;
	} 

	HWND m_hHostMain = (HWND)hWnd;
	CWndFrame* _pFrame = theApp.m_pWndFrame;

	if (b)
		_pFrame->m_bottom = b;
	if (r)
		_pFrame->m_right = r;
	if (l)
		_pFrame->m_left = l;
	if (t)
		_pFrame->m_top = t;
	if (b2)
		_pFrame->m_bottom2 = b2;
	if (r2)
		_pFrame->m_right2 = r2;
	if (l2)
		_pFrame->m_left2 = l2;
	if (t2)
		_pFrame->m_top2 = t2;

	CWnd::FromHandle(m_hHostMain)->ModifyStyle(0, WS_CLIPSIBLINGS | WS_CLIPCHILDREN);

	CWndNode *pRootNode = nullptr;
	theApp.m_pPage = nullptr;
	theApp.m_pCurWorkNodeParse = m_pParse;
	pRootNode = _pFrame->OpenXtmlDocument(theApp.m_strCurrentKey, strXml);
	theApp.m_strCurrentKey = _T("");

	if (pRootNode != nullptr)
	{
		if (bSizable)
		{
			HWND hParent = ::GetParent(pRootNode->m_pHostWnd->m_hWnd);
			CWindow m_wnd;
			m_wnd.Attach(hParent);
			if ((m_wnd.GetStyle() | WS_CHILD) == 0)
			{
				m_wnd.ModifyStyle(0, WS_SIZEBOX | WS_BORDER | WS_MINIMIZEBOX|WS_MAXIMIZEBOX);
			}
			m_wnd.Detach();
			::PostMessage(hParent, WM_TANGRAMMSG, 0, 1965);
		}
	}
	return pRootNode;
}

STDMETHODIMP CTangram::put_ExternalInfo(VARIANT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	switch (newVal.vt)
	{
		case VT_DISPATCH:
		{
			if (theApp.m_pAppDisp == nullptr)
			{
				theApp.m_pAppDisp = newVal.pdispVal;
				theApp.m_pAppDisp->AddRef();
				return S_OK;
			}
		}
		break;
	}

	theApp.m_varApplication = newVal;
	return S_OK;
}

STDMETHODIMP CTangram::get_RootNodes(IWndNodeCollection** pNodeColletion)
{
	if (m_pRootNodes == nullptr)
	{
		CComObject<CWndNodeCollection>::CreateInstance(&m_pRootNodes);
		m_pRootNodes->AddRef();
	}

	m_pRootNodes->m_pNodes->clear();

	for (auto& it : theApp.m_mapWindowPage)
	{
		CWndPage* pFrame = it.second;

		for (auto fit : pFrame->m_mapFrame)
		{
			CWndFrame* pFrame = fit.second;

			for (auto it : pFrame->m_mapNode)
			{
				m_pRootNodes->m_pNodes->push_back(it.second);
			}
		}
	}
	return m_pRootNodes->QueryInterface(IID_IWndNodeCollection, (void**)pNodeColletion);
}

STDMETHODIMP CTangram::get_CurrentActiveNode(IWndNode** pVal)
{
	if (theApp.m_pWndNode)
		*pVal = theApp.m_pWndNode;

	return S_OK;
}

STDMETHODIMP CTangram::GetWndFrame(long hHostWnd, IWndFrame** ppFrame)
{
	HWND m_hHostMain = (HWND)hHostWnd;
	DWORD dwID = ::GetWindowThreadProcessId(m_hHostMain, NULL);
	TangramThreadInfo* pThreadInfo = theApp.GetThreadInfo(dwID);

	CWndFrame* m_pFrame = nullptr;
	auto iter = pThreadInfo->m_mapTangramFrame.find(m_hHostMain);
	if (iter != pThreadInfo->m_mapTangramFrame.end())
	{
		m_pFrame = (CWndFrame*)iter->second;
		*ppFrame = m_pFrame;
	}

	return S_OK;
}

STDMETHODIMP CTangram::SetHostFocus(void)
{
	theApp.m_pWndFrame = nullptr;
	return S_OK;
}

STDMETHODIMP CTangram::CreateCLRObj(BSTR bstrObjID, IDispatch** ppDisp)
{
	theApp.LoadCLR();

	if (theApp.m_pCloudAddinCLRProxy)
	{
		*ppDisp = theApp.m_pCloudAddinCLRProxy->CreateCLRObj(bstrObjID);
		if (*ppDisp)
			(*ppDisp)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CTangram::get_CreatingNode(IWndNode** pVal)
{
	if (theApp.m_pWndNode)
		*pVal = theApp.m_pWndNode;

	return S_OK;
}

STDMETHODIMP CTangram::Encode(BSTR bstrSRC, VARIANT_BOOL bEncode, BSTR* bstrRet)
{
	CString strSRC = OLE2T(bstrSRC);
	if (::PathFileExists(strSRC))
		strSRC = theApp.EncodeFileToBase64(strSRC);
	else if (strSRC != _T(""))
		strSRC = theApp.Encode(strSRC, bEncode ? true : false);

	*bstrRet = strSRC.AllocSysString();
	return S_OK;
}

STDMETHODIMP CTangram::get_Application(IDispatch** pVal)
{
	if (theApp.m_pAppDisp)
	{
		*pVal = theApp.m_pAppDisp;
		(*pVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CTangram::get_AppExtender(BSTR bstrKey, IDispatch** pVal)
{
	CString strName = OLE2T(bstrKey);
	if (strName != _T(""))
	{
		auto it = m_mapObjDic.find(strName);
		if (it == m_mapObjDic.end())
			return S_OK;
		else
		{
			* pVal = it->second;
			(*pVal)->AddRef();
		}
	}

	return S_OK;
}

STDMETHODIMP CTangram::put_AppExtender(BSTR bstrKey, IDispatch* newVal)
{
	CString strName = OLE2T(bstrKey);
	if (strName != _T(""))
	{
		auto it = m_mapObjDic.find(strName);
		if (it != m_mapObjDic.end())
		{
			it->second->Release();
			m_mapObjDic.erase(it);
		}
		if (newVal != nullptr)
		{
			m_mapObjDic[strName] = newVal;
			newVal->AddRef();
		}
#ifndef _WIN64
		if (strName.CompareNoCase(_T("DTE")) == 0)
		{
			VSCloudPlus::VisualStudioPlus::CVSCloudAddin* pAddin = (VSCloudPlus::VisualStudioPlus::CVSCloudAddin*)this;
			CComQIPtr<VxDTE::_DTE> pDTE(newVal);
			pAddin->m_pDTE = pDTE.p;
			pAddin->m_pDTE->AddRef();
			pAddin->OnInitInstance();
		}
#endif
	}
	return S_OK;
}

STDMETHODIMP CTangram::get_AppKeyValue(BSTR bstrKey, VARIANT* pVal)
{
	CString strKey = OLE2T(bstrKey);

	if (strKey != _T(""))
	{
		strKey.Trim();
		strKey.MakeLower();
		auto it = theApp.m_mapValInfo.find(strKey);
		if (it != theApp.m_mapValInfo.end())
		{
			*pVal = it->second;
			return S_OK;
		}
	}
	return S_FALSE;
}

STDMETHODIMP CTangram::put_AppKeyValue(BSTR bstrKey, VARIANT newVal)
{
	CString strKey = OLE2T(bstrKey);

	if (strKey == _T(""))
		return S_OK;
	strKey.Trim();
	strKey.MakeLower();

	auto it = theApp.m_mapValInfo.find(strKey);
	if (it != theApp.m_mapValInfo.end())
	{
		::VariantClear(&it->second);
		theApp.m_mapValInfo.erase(it);
	}
	if (strKey.CompareNoCase(_T("CLRProxy")) == 0)
	{
		if (theApp.m_pCloudAddinCLRProxy == nullptr&&newVal.llVal)
		{
			theApp.m_pCloudAddinCLRProxy = (CApplicationCLRProxyImpl *)newVal.llVal;

			if (theApp.m_hCBTHook == NULL)
				theApp.m_hCBTHook = SetWindowsHookEx(WH_CBT, CTangramApp::CBTProc, NULL, GetCurrentThreadId());
		}
		return S_OK;
	}
	//if (strKey.CompareNoCase(_T("JNIEnv")) == 0)
	//{
	//	//if (newVal.llVal)
	//	//{
	//	//	if (theApp.m_pJVM)
	//	//	{
	//	//		////theApp.m_mapValInfo[strKey] = newVal;
	//	//		////((EclipseCloudPlus::EclipsePlus::CEclipseCloudAddin*)this)->InitTangramforJava();
	//	//		//structured_task_group tasks;
	//	//		//auto task = create_task([&newVal,this]()
	//	//		//{
	//	//		//	((EclipseCloudPlus::EclipsePlus::CEclipseCloudAddin*)this)->InitTangramforJava();
	//	//		//});
	//	//	}
	//	//}
	//	return S_OK;
	//}
	if (strKey.CompareNoCase(_T("JVM")) == 0)
	{
		if (newVal.llVal)
		{
			theApp.m_mapValInfo[strKey] = newVal;
			theApp.m_pJVM = (JavaVM *)newVal.lVal;
		}
		return S_OK;
	}
	if (strKey.CompareNoCase(_T("EclipseStart")) == 0)
	{
		EclipseCloudPlus::EclipsePlus::CEclipseCloudAddin* pAddin = ((EclipseCloudPlus::EclipsePlus::CEclipseCloudAddin*)theApp.m_pHostCore);
		pAddin->InitEclipsePlus();

		return S_OK;
	}
	if (strKey.CompareNoCase(_T("TangramProxy")) == 0)
	{
		if (theApp.m_pTangramProxy == nullptr&&newVal.llVal)
		{
			theApp.m_pTangramProxy = (CTangramProxy *)newVal.llVal;
			theApp.m_bEnableProcessFormTabKey = true;
			if (theApp.m_hCBTHook == NULL)
				theApp.m_hCBTHook = SetWindowsHookEx(WH_CBT, CTangramApp::CBTProc, NULL, GetCurrentThreadId());
		}
		else
		{
			theApp.m_pTangramProxy = nullptr;
			UnhookWindowsHookEx(theApp.m_hCBTHook);
		}
		return S_OK;
	}

	theApp.m_mapValInfo[strKey] = newVal;

	if (strKey.CompareNoCase(_T("TangramDesignerXml")) == 0)
	{
		if (theApp.m_pDocDOMTree)
		{
			CString strXml = OLE2T(newVal.bstrVal);
			theApp.m_pDocDOMTree->DeleteItem(theApp.m_pDocDOMTree->m_hFirstRoot);
			if (theApp.m_pDocDOMTree->m_pHostXmlParse)
			{
				delete theApp.m_pDocDOMTree->m_pHostXmlParse;
			}
			theApp.InitDesignerTreeCtrl(strXml);
		}
		return S_OK;
	}
	if (strKey.CompareNoCase(_T("EnableProcessFormTabKey")) == 0)
	{
		if (newVal.vt == VT_I4&&newVal.lVal == 0)
		{
			theApp.m_bEnableProcessFormTabKey = false;
			return S_OK;
		}
	}
	if (strKey.CompareNoCase(_T("TangramPackageProxy")) == 0)
	{
		if (newVal.llVal)
		{
			theApp.m_pTangramPackageProxy = (CTangramPackageProxy*)newVal.llVal;
			theApp.m_bEnableProcessFormTabKey = false;
			return S_OK;
		}
	}

	if (strKey.CompareNoCase(_T("mainwnd")) == 0)
	{
		if ((newVal.vt == VT_I4 || newVal.vt == VT_I8) && newVal.lVal)
		{
			HWND hWnd = (HWND)newVal.lVal;
			if (::IsWindow(hWnd))
			{
				if (m_hMainWnd == NULL)
				{
					m_hMainWnd = hWnd;
				}
			}
			return S_OK;
		}
	}
	if (strKey.CompareNoCase(_T("defaultframewnd")) == 0)
	{
		if ((newVal.vt == VT_I4 || newVal.vt == VT_I8) && newVal.lVal)
		{
			HWND hWnd = (HWND)newVal.lVal;
			if (::IsWindow(hWnd))
			{
				m_hMainFrameWnd = hWnd;
			}
			return S_OK;
		}
	}
	return S_OK;
}

STDMETHODIMP CTangram::MessageBox(long hWnd, BSTR bstrContext, BSTR bstrCaption, long nStyle, int* nRet)
{
	*nRet = ::MessageBox((HWND)hWnd, OLE2T(bstrContext), OLE2T(bstrCaption), nStyle);
	return S_OK;
}

STDMETHODIMP CTangram::NewGUID(BSTR* retVal)
{
	GUID   m_guid;
	CString   strGUID = _T("");
	if (S_OK == ::CoCreateGuid(&m_guid))
	{
		strGUID.Format(_T("%08X-%04X-%04x-%02X%02X-%02X%02X%02X%02X%02X%02X"),
			m_guid.Data1, m_guid.Data2, m_guid.Data3, m_guid.Data4[0], m_guid.Data4[1], m_guid.Data4[2], m_guid.Data4[3],
			m_guid.Data4[4], m_guid.Data4[5], m_guid.Data4[6], m_guid.Data4[7]);
		*retVal = strGUID.AllocSysString();
	}

	return S_OK;
}

STDMETHODIMP CTangram::TangramGetObject(IDispatch* SourceDisp, IDispatch** ResultDisp)
{
	IStream* pStream = 0;
	HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_IDispatch, SourceDisp, &pStream);
	if (hr == S_OK)
	{
		IDispatch* pEventTarget = nullptr;
		hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_IDispatch, (LPVOID *)&pEventTarget);
		if (hr == S_OK&&pEventTarget)
		{
			*ResultDisp = pEventTarget;
		}
	}
	return S_OK;
}

STDMETHODIMP CTangram::GetCLRControl(IDispatch* CtrlDisp, BSTR bstrNames, IDispatch** ppRetDisp)
{
	CString strNames = OLE2T(bstrNames);
	if (theApp.m_pCloudAddinCLRProxy&&strNames != _T("") && CtrlDisp)
		*ppRetDisp = theApp.m_pCloudAddinCLRProxy->GetCLRControl(CtrlDisp, bstrNames);

	return S_OK;
}

STDMETHODIMP CTangram::ActiveCLRMethod(BSTR bstrObjID, BSTR bstrMethod, BSTR bstrParam, BSTR bstrData)
{
	theApp.LoadCLR();

	if (theApp.m_pCloudAddinCLRProxy)
		theApp.m_pCloudAddinCLRProxy->ActiveCLRMethod(bstrObjID, bstrMethod, bstrParam, bstrData);

	return S_OK;
}

STDMETHODIMP CTangram::CreateWndPage(long hWnd, IWndPage** ppTangram)
{
	HWND _hWnd = (HWND)hWnd;
	if (::IsWindow(_hWnd))
	{
		CWndPage* pPage = nullptr;
		auto it = theApp.m_mapWindowPage.find(_hWnd);
		if (it != theApp.m_mapWindowPage.end())
			pPage = it->second;
		else
		{
			pPage = new CComObject<CWndPage>();
			pPage->m_hWnd = _hWnd;
			theApp.m_mapWindowPage[_hWnd] = pPage;
		}
		if (_hWnd == m_hMainWnd)
		{
			m_mapObjDic[_T("defaulttangram")] = pPage;
		}
		*ppTangram = pPage;
	}
	return S_OK;
}

STDMETHODIMP CTangram::GetItemText(IWndNode* pNode, long nCtrlID, LONG nMaxLengeh, BSTR* bstrRet)
{
	if (pNode == nullptr)
		return S_OK;
	long h = 0;
	pNode->get_Handle(&h);

	HWND hWnd = (HWND)h;
	if (::IsWindow(hWnd))
	{
		if (nMaxLengeh == 0)
		{
			hWnd = ::GetDlgItem(hWnd, nCtrlID);
			m_HelperWnd.Attach(hWnd);
			CString strText(_T(""));
			m_HelperWnd.GetWindowText(strText);
			m_HelperWnd.Detach();
			*bstrRet = strText.AllocSysString();
		}
		else
		{
			LPWSTR lpsz = _T("");
			::GetDlgItemText(hWnd, nCtrlID, lpsz, nMaxLengeh);
			*bstrRet = CComBSTR(lpsz);
		}
	}
	return S_OK;
}

STDMETHODIMP CTangram::SetItemText(IWndNode* pNode, long nCtrlID, BSTR bstrText)
{
	if (pNode == nullptr)
		return S_OK;
	long h = 0;
	pNode->get_Handle(&h);

	HWND hWnd = (HWND)h;
	if (::IsWindow(hWnd))
		::SetDlgItemText(hWnd, nCtrlID, OLE2T(bstrText));

	return S_OK;
}

STDMETHODIMP CTangram::ConnectApp(BSTR bstrAppID, ITangram** ResultTangramCore)
{
	CString strAppID = OLE2T(bstrAppID);
	strAppID.Trim();
	strAppID.MakeLower();
	if (strAppID != _T(""))
	{
		if (theApp.m_mapOfficeAppID.size() == 0)
		{
			theApp.m_mapOfficeAppID[_T("word.application")] = 0;
			theApp.m_mapOfficeAppID[_T("excel.application")] = 1;
			theApp.m_mapOfficeAppID[_T("outlook.application")] = 2;
			theApp.m_mapOfficeAppID[_T("onenote.application")] = 3;
			theApp.m_mapOfficeAppID[_T("infopath.application")] = 4;
			theApp.m_mapOfficeAppID[_T("project.application")] = 5;
			theApp.m_mapOfficeAppID[_T("visio.application")] = 6;
			theApp.m_mapOfficeAppID[_T("access.application")] = 7;
			theApp.m_mapOfficeAppID[_T("powerpoint.application")] = 8;
			theApp.m_mapOfficeAppID[_T("lync.ucofficeintegration.1")] = 9;
		}

		auto it = theApp.m_mapRemoteTangramCore.find(strAppID);
		if (it == theApp.m_mapRemoteTangramCore.end())
		{
			CComPtr<IDispatch> pApp;
			pApp.CoCreateInstance(bstrAppID, 0, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER);
			if (pApp)
			{
				auto it2 = theApp.m_mapOfficeAppID.find(strAppID);
				if (it2 != theApp.m_mapOfficeAppID.end())
				{
					int nIndex = it2->second;
					CComPtr<Office::COMAddIns> pAddins;
					switch (nIndex)
					{
					case 0:
					{
						CComQIPtr<Word::_Application> pWordApp(pApp);
						if (pWordApp)
						{
							pWordApp->put_Visible(true);
							pWordApp->get_COMAddIns(&pAddins);
						}
					}
					break;
					case 1:
					{
						CComQIPtr<Excel::_Application> pExcelApp(pApp);
						pExcelApp->put_UserControl(true);
						pExcelApp->get_COMAddIns(&pAddins);
						pExcelApp->put_Visible(0, true);
					}
					break;
					case 2:
					{
						CComQIPtr<OutLook::_Application> pOutLookApp(pApp);
						pOutLookApp->get_COMAddIns(&pAddins);
					}
					break;
					case 3:
					{
						//CComQIPtr<OneNote::_Application> pOneNoteApp(pApp);
						//pOneNoteApp->get_COMAddIns(&pAddins);
					}
					break;
					case 4:
					{
						//CComQIPtr<OneNote::_Application> pOneNoteApp(pApp);
						//pOneNoteApp->get_COMAddIns(&pAddins);
					}
					break;
					case 5:
					{
						CComQIPtr<MSProject::_MSProject> pProjectApp(pApp);
						//pProjectApp->get_COMAddIns(&pAddins);
					}
					break;
					case 6:
					{
						//CComQIPtr<OneNote::_Application> pOneNoteApp(pApp);
						//pOneNoteApp->get_COMAddIns(&pAddins);
					}
					break;
					case 7:
					{
						//CComQIPtr<OneNote::_Application> pOneNoteApp(pApp);
						//pOneNoteApp->get_COMAddIns(&pAddins);
					}
					break;
					case 8:
					{
						CComQIPtr<PowerPoint::_Application> pPptApp(pApp);
						pPptApp->get_COMAddIns(&pAddins);
					}
					break;
					case 9:
					{
						using namespace UCCollaborationLib;
						CComPtr<IUCOfficeIntegration> _pUCOfficeIntegration;
						IDispatch*			pLyncClient		  = nullptr;
						ILyncClient*		m_pLyncClient	  = nullptr;
						IContactManager*	m_pContactManager = nullptr;
						_pUCOfficeIntegration->GetInterface(CComBSTR(_T("16.0.0.0")), oiInterfaceILyncClient, (IDispatch **)&pLyncClient);
						HRESULT hr = pLyncClient->QueryInterface(IID_ILyncClient, (void**)&m_pLyncClient);
						m_pLyncClient->get_ContactManager(&m_pContactManager);
					}
					break;
					default:
						break;
					}

					if (pAddins)
					{
						CComPtr<Office::COMAddIn> pAddin;
						pAddins->Item(&CComVariant(_T("tangram.tangram")), &pAddin);
						if (pAddin)
						{
							CComPtr<IDispatch> pAddin2;
							pAddin->get_Object(&pAddin2);
							CComQIPtr<ITangram> _pTangramAddin(pAddin2);
							if (_pTangramAddin)
							{
								theApp.m_mapRemoteTangramCore[strAppID] = _pTangramAddin.p;
								_pTangramAddin.p->AddRef();
								long h = 0;
								_pTangramAddin->get_RemoteHelperHWND(&h);
								if (h)
								{
									HWND hWnd = (HWND)h;
									CHelperWnd* pWnd = new CHelperWnd();
									pWnd->m_strID = strAppID;
									pWnd->Create(hWnd, 0, _T(""), WS_CHILD);
									theApp.m_mapRemoteTangramHelperWnd[strAppID] = pWnd;
								}
							}
						}
					}
				}
				else
				{
					DISPID dispID = 0;
					DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
					VARIANT result = { 0 };
					EXCEPINFO excepInfo;
					memset(&excepInfo, 0, sizeof excepInfo);
					UINT nArgErr = (UINT)-1; // initialize to invalid arg
					LPOLESTR func = L"Tangram";
					HRESULT hr = pApp->GetIDsOfNames(GUID_NULL, &func, 1, LOCALE_SYSTEM_DEFAULT, &dispID);
					if (S_OK == hr)
					{
						hr = pApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &result, &excepInfo, &nArgErr);
						if (S_OK == hr && VT_DISPATCH == result.vt)
						{
							CComQIPtr<ITangram> pCloudAddin(result.pdispVal);
							if (pCloudAddin)
							{
								*ResultTangramCore = pCloudAddin.p;
								theApp.m_mapRemoteTangramCore[strAppID] = pCloudAddin.p;

								pCloudAddin.p->AddRef();
								long h = 0;
								pCloudAddin->get_RemoteHelperHWND(&h);
								if (h)
								{
									HWND hWnd = (HWND)h;
									CHelperWnd* pWnd = new CHelperWnd();
									pWnd->m_strID = strAppID;
									pWnd->Create(hWnd, 0, _T(""), WS_CHILD);
									theApp.m_mapRemoteTangramHelperWnd[strAppID] = pWnd;
									m_mapAppDispDic[strAppID] = pApp.Detach();
								}
								pCloudAddin.Detach();
								::VariantClear(&result);
							}
						}
					}
				}
			}
		}
		else
		{
			*ResultTangramCore = it->second;
			(*ResultTangramCore)->AddRef();
		}
	}

	return S_OK;
}

STDMETHODIMP CTangram::CreateOfficeObj(BSTR bstrAppID, BSTR bstrXml)
{
	CString strAppID = OLE2T(bstrAppID);
	strAppID.Trim();
	strAppID.MakeLower();
	auto it = theApp.m_mapRemoteTangramCore.find(strAppID);
	if (it == theApp.m_mapRemoteTangramCore.end())
	{
		ITangram* pCloudAddin = nullptr;
		ConnectApp(strAppID.AllocSysString(),&pCloudAddin);
		it = theApp.m_mapRemoteTangramCore.find(strAppID);
	}
	if (it != theApp.m_mapRemoteTangramCore.end())
		it->second->CreateOfficeDocument(bstrXml);

	return S_OK;
}

STDMETHODIMP CTangram::DownLoadFile(BSTR bstrFileURL, BSTR bstrTargetFile, BSTR bstrActionXml)
{
	CString  strFileURL = OLE2T(bstrFileURL);
	strFileURL.Trim();
	if (theApp.CheckUrl(strFileURL) == false)
		return S_FALSE;
	if (strFileURL == _T(""))
		return S_FALSE;
	CString strTargetFile = OLE2T(bstrTargetFile);
	CString _strTarget = _T("");
	int nPos = strTargetFile.Find(_T("\\"));
	if (nPos != -1)
	{
		_strTarget = strTargetFile.Left(nPos);
		if (_strTarget.CompareNoCase(_T("TangramData")) == 0)
		{
			_strTarget = strTargetFile.Mid(nPos);
			strTargetFile = theApp.m_strAppDataPath + _strTarget;
		}
		else
			_strTarget = _T("");
	}

	nPos = strTargetFile.ReverseFind('\\');
	if (nPos != -1)
	{
		CString strDir = strTargetFile.Left(nPos);
		if (::PathIsDirectory(strDir) == false)
			::SHCreateDirectory(NULL, strDir);
	}

	CloudUtilities::CDownLoadObj* pDownLoadoObj = new CloudUtilities::CDownLoadObj();
	pDownLoadoObj->m_strActionXml = OLE2T(bstrActionXml);
	pDownLoadoObj->DownLoadFile(strFileURL, strTargetFile);
	return S_OK;
}

STDMETHODIMP CTangram::ExtendXml(BSTR bstrXml, BSTR bstrKey, IDispatch** ppNode)
{
	AFX_MANAGE_STATE(AfxGetAppModuleState());
	if (m_hMainWnd == 0)
		return S_FALSE;
	if (m_pPage == nullptr)
	{
		CreateWndPage((long)m_hMainWnd, &m_pPage);
	}
	if (m_pPage)
	{
		if (m_pFrame == nullptr&&m_hMainFrameWnd&&::IsWindow(m_hMainFrameWnd))
		{
			theApp.m_pPage->CreateFrame(CComVariant(0), CComVariant((long)m_hMainFrameWnd), CComBSTR(L"tangram"), &m_pFrame);
		}
		if (m_pFrame)
		{
			CString strXml = OLE2T(bstrXml);
			strXml.Trim();
			auto it = m_strMapKey.find(strXml);
			if (it == m_strMapKey.end())
			{
				m_strMapKey[strXml] = OLE2T(bstrKey);
				IWndNode* pNode = nullptr;
				m_pFrame->Extend(bstrKey, bstrXml, &pNode);
			}
			else
				m_pFrame->Extend(CComBSTR(it->second), bstrXml, (IWndNode**)ppNode);
		}
	}
	::BringWindowToTop(m_hMainWnd);
	::SetFocus(m_hMainWnd);
	return S_OK;
}

STDMETHODIMP CTangram::InitVBAForm(IDispatch* pFormDisp, long nStyle, BSTR bstrXml, IWndNode** ppNode)
{
	return S_OK;
}

STDMETHODIMP CTangram::get_DesignNode(IWndNode** pVal)
{
	if (theApp.m_pDesignWindowNode)
		(*pVal) = (IWndNode*)theApp.m_pDesignWindowNode;
	return S_OK;
}

STDMETHODIMP CTangram::put_CurrentDesignNode(IWndNode* newVal)
{
	if (m_hChildHostWnd == nullptr&&theApp.m_pTangramPackageProxy==nullptr)
		CreateCommonDesignerToolBar();
	theApp.m_pHostDesignUINode = (CWndNode*)newVal;
	if(theApp.m_pDocDOMTree)
		theApp.m_pHostDesignUINode->m_pRootObj->m_pDocTreeCtrl = theApp.m_pDocDOMTree;
	return S_OK;
}

STDMETHODIMP CTangram::UpdateWndNode(IWndNode* pNode)
{
	CWndNode* pWindowNode = (CWndNode*)pNode;
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
			UpdateWndNode(it2);
		}

		if (pWindowNode->m_pOfficeObj)
		{
			CTangramXmlParse* pWndParse = pWindowNode->m_pTangramParse->GetChild(_T("window"));
			CString strXml = pWndParse->xml();
			CString strNodeName = pWindowNode->m_pTangramParse->name();
			theApp.m_pHostCore->UpdateOfficeObj(pWindowNode->m_pOfficeObj, strXml, strNodeName);
		}
	}

	return S_OK;
}

STDMETHODIMP CTangram::GetNodeFromeHandle(long hWnd, IWndNode** ppRetNode)
{
	HWND _hWnd = (HWND)hWnd;
	if (::IsWindow(_hWnd))
	{
		LRESULT lRes = ::SendMessage(_hWnd, WM_TANGRAMGETNODE, 0, 0);
		if (lRes)
		{
			*ppRetNode = (IWndNode*)lRes;
		}
	}
	return S_OK;
}

STDMETHODIMP CTangram::GetCLRControlString(BSTR bstrAssemblyPath, BSTR* bstrCtrls)
{
	CString strPath = OLE2T(bstrAssemblyPath);
	int nPos = strPath.ReverseFind('\\');
	CString strLib = strPath.Mid(nPos + 1);
	nPos = strLib.Find(_T("."));
	strLib = strLib.Left(nPos).MakeLower();
	auto it = theApp.m_mapValInfo.find(strLib);
	theApp.LoadCLR();
	if (it == theApp.m_mapValInfo.end())
	{
		CString strCtrls = theApp.m_pCloudAddinCLRProxy->GetControlStrs(strLib).MakeLower();
		if (strCtrls != _T(""))
		{
			theApp.m_mapValInfo[strLib] = CComVariant(strCtrls);
		}
		*bstrCtrls = strCtrls.AllocSysString();
	}
	else
		*bstrCtrls = it->second.bstrVal;
	return S_OK;
}

STDMETHODIMP CTangram::GetNewLayoutNodeName(BSTR bstrCnnID, IWndNode* pDesignNode, BSTR* bstrNewName)
{
	BOOL bGetNew = false;
	CString strNewName = _T("");
	CString strName = OLE2T(bstrCnnID);
	int nIndex = 0;
	CWndNode* _pNode = ((CWndNode*)pDesignNode);
	CWndNode* pNode = _pNode->m_pRootObj;
	auto it = pNode->m_mapLayoutNodes.find(strName);
	if (it == pNode->m_mapLayoutNodes.end())
	{
		*bstrNewName = strName.AllocSysString();
		return S_OK;
	}
	while (bGetNew == false)
	{
		strNewName.Format(_T("%s%d"),strName, nIndex);
		it = pNode->m_mapLayoutNodes.find(strNewName);
		if (it == pNode->m_mapLayoutNodes.end())
		{
			*bstrNewName = strNewName.AllocSysString();
			bGetNew = true;
			return S_OK;
		}
		nIndex++;
	}

	return S_OK;
}

STDMETHODIMP CTangram::CreateWebTask(BSTR bstrTaskName, BSTR bstrTaskURL, IWebTaskObj** pWebTaskObj)
{
	CString strName = OLE2T(bstrTaskName);
	if (strName != _T(""))
	{
		auto it = m_mapWebTask.find(strName);
		if (it != m_mapWebTask.end())
		{
			*pWebTaskObj = (IWebTaskObj*)it->second;
		}
		else
		{
			theApp.Lock();
			CWebTaskObj* pTaskObj = new CComObject<CWebTaskObj>;
			pTaskObj->m_strName = strName;
			m_mapWebTask[strName] = pTaskObj;
			pTaskObj->AddRef();
			*pWebTaskObj = (IWebTaskObj*)pTaskObj;
			theApp.Unlock();
		}
		(*pWebTaskObj)->AddRef();
	}

	return S_OK;
}

STDMETHODIMP CTangram::CreateTaskObj(ITaskObj** ppTaskObj)
{
	CTaskObj* pTaskObj = new CComObject<CTaskObj>;
	pTaskObj->AddRef();
	*ppTaskObj = pTaskObj;

	return S_OK;
}

STDMETHODIMP CTangram::CreateRestObj(IRestObject** ppRestObj)
{
	CRestObj* pRestObj = new CComObject<CRestObj>;
	pRestObj->AddRef();
	*ppRestObj = pRestObj;

	return S_OK;
}


CString CTangram::GetDocTemplateXml(CString strCaption, CString _strPath)
{
	CString strTemplate = _T("");

	theApp.Lock();
	auto it = theApp.m_mapValInfo.find(_T("doctemplate"));
	if (it != theApp.m_mapValInfo.end())
	{
		strTemplate = OLE2T(it->second.bstrVal);
		::VariantClear(&it->second);
		theApp.m_mapValInfo.erase(it);
	}
	if (strTemplate == _T(""))
	{
		CDocTemplateDlg dlg;
		CString strPath = _T("DocTemplate\\");
		if (_strPath != _T(""))
		{
			strPath += _strPath;
			strPath += _T("\\");
		}
		dlg.m_strDir = strPath;
		dlg.m_strCaption = strCaption;
		if (dlg.DoModal() == IDOK)
			strTemplate = dlg.m_strDocTemplatePath;
		else
			strTemplate = m_strDefaultTemplateXml;
	}
	theApp.Unlock();
	return strTemplate;
}

STDMETHODIMP CTangram::AddVBAFormsScript(IDispatch* OfficeObject, BSTR bstrKey, BSTR bstrXml)
{
	return S_OK;
}

STDMETHODIMP CTangram::ExportOfficeObjXml(IDispatch* OfficeObject, BSTR* bstrXml)
{
	return S_OK;
}

STDMETHODIMP CTangram::GetFrameFromVBAForm(IDispatch* pForm, IWndFrame** ppFrame)
{
	return S_OK;
}

STDMETHODIMP CTangram::GetActiveTopWndNode(IDispatch* pForm, IWndNode** WndNode)
{
	return S_OK;
}

STDMETHODIMP CTangram::GetObjectFromWnd(LONG hWnd, IDispatch** ppObjFromWnd)
{
	return S_OK;
}


STDMETHODIMP CTangram::ReleaseTangram()
{
	theApp.m_pHostCore = nullptr;
	theApp.m_pTangram = nullptr;
	delete this;

	return S_OK;
}

STDMETHODIMP CTangram::put_JavaProxy(IJavaProxy* newVal)
{
	if (theApp.m_pJavaProxy == nullptr)
	{
		theApp.m_pJavaProxy = newVal;
		theApp.m_pJavaProxy->AddRef();
	}

	return S_OK;
}

STDMETHODIMP CTangram::InitJava(int nIndex)
{
	theApp.InitJava(nIndex);
	return S_OK;
}


STDMETHODIMP CTangram::get_HostWnd(LONG* pVal)
{
	*pVal = (LONG)m_hHostWnd;

	return S_OK;
}


STDMETHODIMP CTangram::put_CollaborationProxy(ICollaborationProxy* newVal)
{
	return S_OK;
}

