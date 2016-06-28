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
#include "msppt.h"
#include "../CloudAddin.h"
#include "PowerpointPlusEvents.h"

using namespace PowerPoint;
using namespace OfficeCloudPlus::PowerPointPlus;
using namespace OfficeCloudPlus::PowerPointPlus::PowerPointPlusEvent;

namespace OfficeCloudPlus
{
	namespace PowerPointPlus
	{
		class CCloudAddinPresentation;
		class CPowerPntObject : public COfficeObject
		{
		public:
			CPowerPntObject(void);
			~CPowerPntObject(void);

			CCloudAddinPresentation*	m_pPresentation;
			void OnObjDestory();
		};

		class CCloudAddinPresentation : 
			public CTangramDocument,
			public map<HWND, CPowerPntObject*>
		{
		public:
			CCloudAddinPresentation();
			~CCloudAddinPresentation();
			CString						m_strKey;
		};

		class CPowerPntCloudAddin : 
			public CCloudAddin,
			public CPowerpointPlusEApplication
		{
		public:
			CPowerPntCloudAddin();
			CPowerPntObject*						m_pActivePowerPntObject;
			CCloudAddinPresentation*				m_pCurrentSavingPresentation;
			map<CString, CCloudAddinPresentation*>	m_mapTangramPres;

		private:
			CComPtr<PowerPoint::_Application>		m_pApplication;

			//CTangramApp:
			STDMETHOD(TangramCommand)(IDispatch* RibbonControl);
			HRESULT OnConnection(IDispatch* pHostApp, int ConnectMode);
			HRESULT OnDisconnection(int DisConnectMode);

			void OnPresentationActivate(CPowerPntObject*);
			void __stdcall OnPresentationSave(/*[in]*/ _Presentation * Pres);
			void __stdcall OnPresentationBeforeSave(/*[in]*/ _Presentation * Pres,	/*[in,out]*/ VARIANT_BOOL * Cancel);

			STDMETHOD(AddDocXml)(IDispatch* pDocdisp, BSTR bstrXml, BSTR bstrKey, VARIANT_BOOL* bSuccess);
			STDMETHOD(GetDocXmlByKey)(IDispatch* pDocdisp, BSTR bstrKey, BSTR* pbstrXml);
			STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);
			
			void UpdateOfficeObj(IDispatch* pObj, CString strXml, CString strName);
			void WindowCreated(LPCTSTR strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd);
			void WindowDestroy(HWND hWnd);
			void ConnectOfficeObj(HWND hWnd);
			bool OnActiveOfficeObj(HWND hWnd);
		};
	}
}// namespace OfficeCloudPlus


