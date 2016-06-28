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

#pragma once
//#include "dllmain.h"

#include "TangramObj.h"
#include "Tangram.h"       // main symbols
using namespace System;
using namespace System::Reflection;

using namespace TangramCLR;

class CTangramCoreCLREvent 
{
public:
	CTangramCoreCLREvent();
	virtual ~CTangramCoreCLREvent();

	void OnClose();
	//void OnDocumentComplete(IDispatch* pDocdisp, BSTR bstrUrl);
	void OnExtendComplete(long hWnd, BSTR bstrUrl, IWndNode* pRootNode);
};

class CWndNodeCLREvent 
{
public:
	CWndNodeCLREvent();
	virtual ~CWndNodeCLREvent();

	gcroot<TangramCLR::WndNode^>	m_pWndNode;

	void OnDestroy();
	void OnNodeAddInsCreated();
	void OnExtendComplete(IWndNode* pNode);
	void OnDocumentComplete(IDispatch* pDocdisp, BSTR bstrUrl);
	void OnNodeAddInCreated(IDispatch* pAddIndisp, BSTR bstrAddInID, BSTR bstrAddInXml);
};
