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
// TangramAddin.cpp : Declaration of the CCloudAddin 

#include "../stdafx.h"
#include "../CloudAddinApp.h"
#include "../WndFrame.h"
#include "../WndNode.h"
#include "../TangramHtmlTreeWnd.h"
#include "../CloudUtilities/CloudAddinDownLoad.h"
#include "../VisualStudioPlus/CloudAddinVSAddin.h"
#include "fm20.h"
#include "CloudAddin.h"
#include "excelplus\excel.h"
#include "wordplus\msword.h"
#include "outlookplus\msoutl.h"
using namespace VBIDE;
/*
Private Sub MDIForm_Load()
Set TangramCore = CreateObject("tangram.tangram")
Dim tangram As Object
Set tangram = TangramCore.CreateWndPage(Me.hWnd)
Dim frameX As Object
Set frameX = tangram.CreateFrame(0, 0, "test")
frameX.Extend "", "d:\AppDoc1.APPXml"
End Sub
*/
namespace OfficeCloudPlus
{
	// CCloudAddin
	CCloudAddin::CCloudAddin()
	{
		if (theApp.m_hCBTHook == NULL)
		{
			theApp.m_hCBTHook = SetWindowsHookEx(WH_CBT, CTangramApp::CBTProc, NULL, GetCurrentThreadId());
		}
		m_hHostWnd = NULL;
		m_strLib = _T("");
		m_strUser = _T("");
		m_strRibbonXml = _T("");
		m_strTemplateXml = _T("");
		m_pPage = nullptr;
		theApp.m_bOfficeApp = true;
		WNDCLASSEX wcex;
		wcex.cbSize = sizeof(WNDCLASSEX);
		wcex.style = CS_HREDRAW | CS_VREDRAW;
		wcex.lpfnWndProc = ::DefWindowProc;
		wcex.cbClsExtra = 0;
		wcex.cbWndExtra = 0;
		wcex.hInstance = theApp.m_hInstance;
		wcex.hIcon = NULL;
		wcex.hCursor = NULL;
		wcex.hbrBackground = NULL;
		wcex.lpszMenuName = NULL;
		wcex.lpszClassName = L"Tangram Remote Helper Window";
		wcex.hIconSm = NULL;
		RegisterClassEx(&wcex);
	}

	CCloudAddin::~CCloudAddin()
	{
		m_strLib = _T("");
		ATLTRACE(_T("**********CCloudAddin::~CCloudAddin: %x*********************\n"), this);
	}

	STDMETHODIMP CCloudAddin::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** /*custom*/)
	{
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::OnDisconnection(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY ** /*custom*/)
	{
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::OnAddInsUpdate(SAFEARRAY ** /*custom*/)
	{
		OnUpdate();
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::OnStartupComplete(SAFEARRAY ** /*custom*/)
	{
		StartupComplete();
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::OnBeginShutdown(SAFEARRAY ** /*custom*/)
	{
		return S_OK;
	}

	void CCloudAddin::OnCloseOfficeObj(CString strName, HWND hWnd)
	{

	}

	void CCloudAddin::_AddDocXml(Office::_CustomXMLParts* pCustomXMLParts, BSTR bstrXml, BSTR bstrKey, VARIANT_BOOL* bSuccess)
	{
		*bSuccess = false;
	}

	void CCloudAddin::_GetDocXmlByKey(Office::_CustomXMLParts* pCustomXMLParts, BSTR bstrKey, BSTR* pbstrXml)
	{

	}

	STDMETHODIMP CCloudAddin::InitVBAForm(IDispatch* pFormDisp, long nStyle, BSTR bstrXml, IWndNode** ppNode)
	{
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::AddVBAFormsScript(IDispatch* OfficeObject, BSTR bstrKey, BSTR bstrXml)
	{
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::ExportOfficeObjXml(IDispatch* OfficeObject, BSTR* bstrXml)
	{
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::GetFrameFromVBAForm(IDispatch* pForm, IWndFrame** ppFrame)
	{
		CComQIPtr<IOleWindow> pOleWnd(pForm);
		if (pOleWnd)
		{
			HWND hWnd = NULL;
			pOleWnd->GetWindow(&hWnd);
			auto it = m_mapVBAForm.find(hWnd);
			if (it != m_mapVBAForm.end())
			{
				*ppFrame = it->second;
			}
		}

		return S_OK;
	}

	STDMETHODIMP CCloudAddin::GetActiveTopWndNode(IDispatch* pForm, IWndNode** WndNode)
	{
		CComQIPtr<IOleWindow> pOleWnd(pForm);
		if (pOleWnd)
		{
			HWND hWnd = NULL;
			pOleWnd->GetWindow(&hWnd);
			auto it = m_mapVBAForm.find(hWnd);
			if (it != m_mapVBAForm.end())
			{
				*WndNode = it->second->m_pWorkNode;
			}
		}
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::GetObjectFromWnd(LONG hWnd, IDispatch** ppObjFromWnd)
	{
		return S_OK;
	}

	// ICustomTaskPaneConsumer Methods
	STDMETHODIMP CCloudAddin::CTPFactoryAvailable(ICTPFactory * CTPFactoryInst)
	{
		CTPFactoryInst->QueryInterface(Office::IID_ICTPFactory, (void**)&m_pCTPFactory);
		return S_OK;
	};


	STDMETHODIMP CCloudAddin::GetCustomUI(BSTR RibbonID, BSTR * RibbonXml)
	{
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::TangramCommand(IDispatch* RibbonControl)
	{
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::TangramOnLoad(IDispatch* RibbonControl)
	{
		m_spRibbonUI = RibbonControl;
		Tangram_OnLoad(RibbonControl);
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::TangramGetImage(BSTR strValue, IPictureDisp ** ppDispImage)
	{
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::TangramGetItemCount(IDispatch* RibbonControl, long* nCount)
	{
		Tangram_GetItemCount(RibbonControl, nCount);
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::TangramGetItemLabel(IDispatch* RibbonControl, long nIndex, BSTR* bstrLabel)
	{
		Tangram_GetItemLabel(RibbonControl, nIndex, bstrLabel);
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::TangramGetItemID(IDispatch* RibbonControl, long nIndex, BSTR* bstrID)
	{
		Tangram_GetItemID(RibbonControl, nIndex, bstrID);
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::TangramGetVisible(IDispatch* RibbonControl, VARIANT* varVisible)
	{
		return S_OK;
	}

	STDMETHODIMP CCloudAddin::get_RemoteHelperHWND(long* pVal)
	{
		if (::IsWindow(m_hHostWnd))
			*pVal = (long)m_hHostWnd;
		return S_OK;
	}

	COfficeObject::COfficeObject(void)
	{
		m_bMDIClient = false;
		m_hClient = NULL;
		m_hTaskPane = NULL;
		m_hTaskPaneWnd = NULL;
		m_hTaskPaneChildWnd = NULL;
		m_pPage = nullptr;
		m_pFrame = nullptr;
		m_pOfficeObj = nullptr;
	}

	COfficeObject::~COfficeObject(void)
	{
	}

	STDMETHODIMP COfficeObject::Show(VARIANT_BOOL bShow)
	{
		if (m_pPage == nullptr)
			return S_OK;
		if (bShow)
		{
			HWND h = ::GetParent(m_hForm);
			m_pFrame->ModifyHost((long)h);
		}
		else
		{
			m_pFrame->ModifyHost((long)m_hChildClient);
		}

		return S_OK;
	}

	STDMETHODIMP COfficeObject::Extend(BSTR bstrKey, BSTR bstrXml, IWndNode** ppNode)
	{
		CString strKey = OLE2T(bstrKey);
		strKey.Trim();
		if (strKey == _T(""))
			return S_FALSE;

		strKey = OLE2T(bstrXml);
		strKey.Trim();
		if (strKey == _T(""))
			return S_FALSE;

		if (m_pPage == nullptr/*&&m_bMDIClient == false*/)
		{
			m_hClient = ::CreateWindowEx(NULL, L"Tangram Remote Helper Window", _T("Tangram Office Plus Addin Helper Window"), WS_CHILD, 0, 0, 0, 0, (HWND)m_hForm, NULL, theApp.m_hInstance, NULL);
			m_hChildClient = ::CreateWindowEx(NULL, L"Tangram Remote Helper Window", _T("Tangram Excel Helper Window"), WS_CHILD, 0, 0, 0, 0, (HWND)m_hClient, NULL, AfxGetInstanceHandle(), NULL);
			theApp.m_pTangram->CreateWndPage((long)m_hClient, &m_pPage);
			if (m_pPage == nullptr)
				return S_FALSE;
			if (m_pFrame == nullptr)
			{
				m_pPage->CreateFrame(CComVariant(0), CComVariant((long)m_hChildClient), CComBSTR(L"Default"), &m_pFrame);
				if (m_pFrame == nullptr)
				{
					delete m_pPage;
					m_pPage = nullptr;
				}
			}
		}

		HRESULT hr = m_pFrame->Extend(bstrKey, bstrXml, ppNode);
		return hr;
	}

	STDMETHODIMP COfficeObject::UnloadTangram()
	{
		return S_OK;
	}
}



