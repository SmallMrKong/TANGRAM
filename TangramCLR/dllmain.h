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

// dllmain.h : Declaration of module class.

#include "TangramCoreEvents.h"
#ifdef _WIN64
#include <ActiveDS.h>
#endif
#include <cstring>
#include <iostream>

//#define WM_USER_TANGRAMTASK	((WM_USER) + 1024)
#define WM_TANGRAMNOTIFY	((WM_USER) + 1965)
#define WM_TANGRAMRECYCLE	((WM_USER) + 1963)
#define WM_TANGRAMRESTATEREDAY	((WM_USER) + 1992)
#define WM_TANGRAMRESTATEQUIT	((WM_USER) + 1993)

class CWndNodeCLREvent;
class CTangramCoreCLREvent;

class CCloudAddinCLRApp : public CAtlDllModuleT< CCloudAddinCLRApp >,public CTangramCoreEvents
{
public:
	CCloudAddinCLRApp();
	~CCloudAddinCLRApp();

	int						m_nAppEndPointCount;
	ITangram*				m_pTangram;
	CTangramCoreCLREvent*	m_pCloudAddinCLREvent;

	bool									m_bHostObjCreated;
	TangramMsgMap							m_mapTangramMsg;

	CString									m_strAppEndpointsScript;

	CRITICAL_SECTION						m_csTaskRecycleSection;
	CRITICAL_SECTION						m_csTaskListSection;

	int CalculateByteMD5(BYTE* pBuffer, int BufferSize, CString &MD5);

private:
	//CTangramCoreEvents:
	void __stdcall OnClose();
	void __stdcall OnExtendComplete(long hWnd, BSTR bstrUrl, IWndNode* pRootNode);
};

extern CCloudAddinCLRApp theApp;

class CTangramNodeEvent : public CWndNodeEvents
{
public:
	CTangramNodeEvent();
	virtual ~CTangramNodeEvent();

	CWndNodeCLREvent* m_pTangramNodeCLREvent;
private:
	void __stdcall  OnExtendComplete();
	void __stdcall  OnDestroy();
	void __stdcall  OnDocumentComplete(IDispatch* pDocdisp, BSTR bstrUrl);
	void __stdcall  OnNodeAddInCreated(IDispatch* pAddIndisp, BSTR bstrAddInID, BSTR bstrAddInXml);
	void __stdcall  OnNodeAddInsCreated();
};
