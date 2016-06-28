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
#include "msprj.h"
#include "../CloudAddin.h"
#include "ProjectPlusEvents.h"
using namespace MSProject;
using namespace OfficeCloudPlus::ProjectPlus::ProjectPlusEvent;

namespace OfficeCloudPlus
{
	namespace ProjectPlus
	{
		class CProjectCloudAddin :
			public CCloudAddin,
			public CTangramEProjectAppEvents,
			public CTangramEProjectAppEvents2
		{
		public:
			CProjectCloudAddin();
			virtual ~CProjectCloudAddin();

			CComPtr<_MSProject> m_pApplication;
			//CTangramApp:
			STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);
			STDMETHOD(TangramCommand)(IDispatch* RibbonControl);
			HRESULT OnConnection(IDispatch* pHostApp, int ConnectMode);
			HRESULT OnDisconnection(int DisConnectMode);
			void WindowCreated(LPCTSTR strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd);
	};
	}
}// namespace OfficeCloudPlus

