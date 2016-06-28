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

#include "def.h"
#include "msacc.h"
#include "AccessEvents.h"
#include "../CloudAddin.h"

using namespace MSAccess;
using namespace OfficeCloudPlus::AccessPlus::AccessPlusEvent;

namespace OfficeCloudPlus
{
	namespace AccessPlus
	{
		class CAccessCloudAddin :
			public CCloudAddin
		{
		public:
			CAccessCloudAddin();
			virtual ~CAccessCloudAddin();

			CComPtr<MSAccess::_Application>			m_pApplication;
		private:
			//CTangramApp:
			STDMETHOD(GetObjectFromWnd)(LONG hWnd, IDispatch** ppObjFromWnd);
			STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);
			STDMETHOD(TangramCommand)(IDispatch* RibbonControl);
			HRESULT OnConnection(IDispatch* pHostApp, int ConnectMode);
			HRESULT OnDisconnection(int DisConnectMode);
			void WindowCreated(LPCTSTR strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd);
			void ConnectOfficeObj(HWND hWnd);
			void WindowDestroy(HWND hWnd);
		};

		class CAccessForm :
			public COfficeObject,
			public CMSAccessFormEvents
		{
		public:
			CAccessForm(void);
			~CAccessForm(void);
			
			void OnObjDestory();
			
			//void __stdcall Load();
			//void __stdcall Current();
			//void __stdcall BeforeInsert(short* Cancel);
			//void __stdcall AfterInsert();
			//void __stdcall BeforeUpdate(short* Cancel);
			//void __stdcall AfterUpdate();
			//void __stdcall Delete(short* Cancel);
			//void __stdcall BeforeDelConfirm(short* Cancel, short* Response);
			//void __stdcall AfterDelConfirm(short* Status);
			//void __stdcall Open(short* Cancel);
			//void __stdcall Resize();
			//void __stdcall Unload(short* Cancel);
			//void __stdcall Close();
			//void __stdcall Activate();
			//void __stdcall Deactivate();
			//void __stdcall GotFocus();
			//void __stdcall LostFocus();
		};
	}
}// namespace OfficeCloudPlus


