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
// CollaborationProxy.h : Declaration of the CCollaborationProxy

#pragma once
#include "resource.h"       // main symbols



#include "Tangram.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


// CCollaborationProxy

class ATL_NO_VTABLE CCollaborationProxy :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCollaborationProxy, &CLSID_CollaborationProxy>,
	public IDispatchImpl<ICollaborationProxy, &IID_ICollaborationProxy, &LIBID_Tangram, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CCollaborationProxy()
	{
	}

BEGIN_COM_MAP(CCollaborationProxy)
	COM_INTERFACE_ENTRY(ICollaborationProxy)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

	STDMETHOD(Start)(BSTR bstrXml);
	STDMETHOD(Close)(BSTR bstrXml);
};
