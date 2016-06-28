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
#pragma once
#include "dte.h"
#include "dte80.h"
#include "def.h"

class CAppDocPackage;

class CTangramApp :
	public CAtlDllModuleT<CTangramApp>
{
public:
	CTangramApp()
	{
		m_bInit = false;
		m_pOutputWindowPane = NULL;
	}

	CString					m_strPath;
	CString					m_strExeName;
	CString					m_strVersion;
	CString					m_strModulePath;
	CString					m_strAppDataPath;
	CString					m_strProgramFilePath;
	CString					m_strWindowSystemPath;
	CString					m_strDesignerXml;

	VxDTE::_DTE*			m_pDTE;
	CAppDocPackage*			m_pTangramPackage;
	OutputWindowPane*		m_pOutputWindowPane;
	CComPtr<ITangram>		m_pTangram;

	BOOL InitInstance();
	bool CheckUrl(CString& url);
	BOOL IsUserAdministrator();
	DWORD ExecCmd(const CString cmd, const BOOL setCurrentDirectory);
private:
	bool					m_bInit;
};

extern CTangramApp theApp;
