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
* OUTLINED IN THE TANGRAM LICENSE AGREEMENT.TANGRAM TEAM
* GRANTS TO YOU (ONE SOFTWARE DEVELOPER) THE LIMITED RIGHT TO USE
* THIS SOFTWARE ON A SINGLE COMPUTER.
*
* CONTACT INFORMATION:
* mailto:sunhuizlz@yeah.net
* http://www.CloudAddin.com
*
********************************************************************************/


#pragma once
#include "../CloudAddin.h"
#include "WordPlusEvents.h"
using namespace OfficeCloudPlus::WordPlus::WordPlusEvent;

namespace OfficeCloudPlus
{
	namespace WordPlus
	{
		class CWordObject;
		class CWordDocument;
		class CWordCloudAddin :
			public CCloudAddin,
			public CWordAppEvents2,
			public map<_Document*, CWordDocument*>
		{
		public:
			CWordCloudAddin();

			CWordObject*				m_pActiveWordObject;
		private:
			CComPtr<_Application>			m_pApplication;

			void __stdcall OnDocumentBeforeSave(_Document* Doc, VARIANT_BOOL* SaveAsUI, VARIANT_BOOL* Cancel);
			CString GetFormXml(CString strFormName);
			void UpdateOfficeObj(IDispatch* pObj, CString strXml, CString strName);

			STDMETHOD(GetDocXmlByKey)(IDispatch* pDocdisp, BSTR bstrKey, BSTR* pbstrXml);
			STDMETHOD(AddDocXml)(IDispatch* pDocdisp, BSTR bstrXml, BSTR bstrKey, VARIANT_BOOL* bSuccess);
			STDMETHOD(CreateOfficeDocument)(BSTR bstrXml);
			STDMETHOD(ExportOfficeObjXml)(IDispatch* OfficeObject, BSTR* bstrXml);
			STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);
			STDMETHOD(TangramCommand)(IDispatch* RibbonControl);

			HRESULT OnConnection(IDispatch* pHostApp, int ConnectMode);
			HRESULT OnDisconnection(int DisConnectMode);
			void WindowCreated(LPCTSTR strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd);
			void WindowDestroy(HWND hWnd);

			void OnOpenDoc(WPARAM);
			void ConnectOfficeObj(HWND hWnd);
			bool OnActiveOfficeObj(HWND hWnd);
			void OnDocActivate(CWordObject*);
		};
	}
}

