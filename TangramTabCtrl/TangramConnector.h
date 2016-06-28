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
* http://www.cloudaddin.com
*
********************************************************************************/
// TangramConnector.h : CTangramConnector implementation file

#pragma once
#include "resource.h"       
#include "TangramTabCtrl_i.h"
#include "Tangram.h"
#include "TangramCategory.h"

// CTangramConnector

class ATL_NO_VTABLE CTangramConnector :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CTangramConnector, &CLSID_TangramConnector>,
	public IDispatchImpl<ICreator, &__uuidof(ICreator)>
{
public:
	CTangramConnector()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_TANGRAMCONNECTOR)


BEGIN_COM_MAP(CTangramConnector)
	COM_INTERFACE_ENTRY(ICreator)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

BEGIN_CATEGORY_MAP(CTangramConnector)
	IMPLEMENTED_CATEGORY(CATID_CloudUIContainerCategory)
END_CATEGORY_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:
	STDMETHOD(Create)(long hParent, IWndNode* pNode, long* hWnd);
	STDMETHOD(get_Names)(BSTR* pVal);
	STDMETHOD(get_Tags)(BSTR bstrName, BSTR* pVal);
};

OBJECT_ENTRY_AUTO(__uuidof(TangramConnector), CTangramConnector)
