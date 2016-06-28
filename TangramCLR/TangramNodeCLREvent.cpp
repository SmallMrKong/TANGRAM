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

#include "stdafx.h"
#include "TangramNodeCLREvent.h"

using namespace System::Runtime::InteropServices;

CTangramCoreCLREvent::CTangramCoreCLREvent()
{
}

CTangramCoreCLREvent::~CTangramCoreCLREvent()
{
}
//
//void CTangramCoreCLREvent::OnDocumentComplete(IDispatch* pDocdisp, BSTR bstrUrl)
//{
//	TangramCLR::TangramCore^ pManager = TangramCLR::TangramCore::GetTangram();
//	pManager->Fire_OnDocumentComplete(Marshal::GetObjectForIUnknown((IntPtr)pDocdisp), BSTR2STRING(bstrUrl));
//}

void CTangramCoreCLREvent::OnExtendComplete(long hWnd, BSTR bstrUrl, IWndNode* pRootNode)
{
	TangramCLR::Tangram^ pManager = TangramCLR::Tangram::GetTangram();
	WndNode^ _pRootNode = theAppProxy._createObject<IWndNode, WndNode>(pRootNode);
	IntPtr nHandle = (IntPtr)hWnd;
	pManager->Fire_OnExtendComplete(nHandle, BSTR2STRING(bstrUrl), _pRootNode);
}


void CTangramCoreCLREvent::OnClose()
{
	TangramCLR::Tangram::GetTangram()->Fire_OnClose();
}

CWndNodeCLREvent::CWndNodeCLREvent()
{
}


CWndNodeCLREvent::~CWndNodeCLREvent()
{
}

void CWndNodeCLREvent::OnExtendComplete(IWndNode* pNode)
{
	m_pWndNode->Fire_ExtendComplete(m_pWndNode);
}

void CWndNodeCLREvent::OnDocumentComplete(IDispatch* pDocdisp, BSTR bstrUrl)
{
	Object^ pObj = reinterpret_cast<Object^>(Marshal::GetObjectForIUnknown((System::IntPtr)(pDocdisp)));
	m_pWndNode->Fire_OnDocumentComplete(m_pWndNode, pObj, BSTR2STRING(bstrUrl));
}

void CWndNodeCLREvent::OnDestroy()
{
	m_pWndNode->Fire_OnDestroy(m_pWndNode);
}

void CWndNodeCLREvent::OnNodeAddInCreated(IDispatch* pAddIndisp, BSTR bstrAddInID, BSTR bstrAddInXml)
{
	Object^ pAddinObj = reinterpret_cast<Object^>(Marshal::GetObjectForIUnknown((System::IntPtr)(pAddIndisp)));
	m_pWndNode->Fire_NodeAddInCreated(m_pWndNode, pAddinObj, BSTR2STRING(bstrAddInID), BSTR2STRING(bstrAddInXml));
}

void CWndNodeCLREvent::OnNodeAddInsCreated()
{
	m_pWndNode->Fire_NodeAddInsCreated(m_pWndNode);
}
