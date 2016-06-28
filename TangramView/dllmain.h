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
* http://www.cloudaddin.com
*
********************************************************************************/
// dllmain.h : 
#include "tangram.h"
#include "TangramView_i.h"
class CCloudAddinDesignerApp : public CWinApp,
	public ATL::CAtlDllModuleT< CCloudAddinDesignerApp >
{
public :
	CString					m_strCtrls;
	CString					m_strLayouts;
	CString					m_strComponents;
	ITangram*				m_pTangram;
	map<CString, CString>	m_mapAssemblyInfo;
	map<CString, CString>	m_mapName;
	map<CString, CString>	m_mapStyle;

	DECLARE_LIBID(LIBID_TangramViewLib)
	BOOL InitInstance();
};

extern class CCloudAddinDesignerApp theApp;
