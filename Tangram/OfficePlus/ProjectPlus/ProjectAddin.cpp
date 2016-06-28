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
********************************************************************************/

#include "../../stdafx.h"
#include "../../CloudAddinApp.h"
#include "ProjectAddin.h"
#include "ProjectPlusEvents.cpp"

namespace OfficeCloudPlus
{
	namespace ProjectPlus
	{
		CProjectCloudAddin::CProjectCloudAddin():CCloudAddin()
		{
		}

		CProjectCloudAddin::~CProjectCloudAddin()
		{
		}

		STDMETHODIMP CProjectCloudAddin::TangramCommand(IDispatch* RibbonControl)
		{
			if (m_spRibbonUI)
				m_spRibbonUI->Invalidate();
			return S_OK;
		}

		HRESULT CProjectCloudAddin::OnConnection(IDispatch* pHostApp, int ConnectMode)
		{
			if (m_strTemplateXml == _T(""))
			{
				CTangramXmlParse m_Parse;
				if (m_Parse.LoadXml(theApp.m_strConfigFile))
				{
					m_strRibbonXml = m_Parse[_T("RibbonUI")][_T("customUI")].xml();
					m_strDefaultTemplateXml = m_Parse[_T("DefaultTemplate")][_T("Tangrams")].xml();
				}
			}
			pHostApp->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pApplication);
			return S_OK;
		}

		HRESULT CProjectCloudAddin::OnDisconnection(int DisConnectMode)
		{
			m_pApplication.p->Release();
			m_pApplication.Detach();
			return S_OK;
		}

		STDMETHODIMP CProjectCloudAddin::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml)
		{
			if (!RibbonXml)
				return E_POINTER;
			*RibbonXml = m_strRibbonXml.AllocSysString();
			return (*RibbonXml ? S_OK : E_OUTOFMEMORY);
		}

		void CProjectCloudAddin::WindowCreated(LPCTSTR lpszClass, LPCTSTR strName, HWND hPWnd, HWND hWnd)
		{
			CString strClassName = lpszClass;
			if (strClassName == _T("OForm"))
			{
			}
		}
	}
}// namespace OfficeCloudPlus

