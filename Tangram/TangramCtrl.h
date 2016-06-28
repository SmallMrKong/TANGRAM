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
********************************************************************************/

// TangramCtrl.h : Declaration of the CTangramCtrl
#pragma once
#include "resource.h"       // main symbols
#include <atlctl.h>
#include "Tangram.h"
#include "EclipsePlus\EclipseAddin.h"
using namespace EclipseCloudPlus::EclipsePlus;

// CTangramCtrl
class ATL_NO_VTABLE CTangramCtrl :
	public CComControl<CTangramCtrl>,
	public CComObjectRootEx<CComSingleThreadModel>,
	public IOleObjectImpl<CTangramCtrl>,
	public IOleInPlaceActiveObjectImpl<CTangramCtrl>,
	public IViewObjectExImpl<CTangramCtrl>,
	public IOleInPlaceObjectWindowlessImpl<CTangramCtrl>,
	public IPersistStreamInitImpl<CTangramCtrl>,
	public IPersistStorageImpl<CTangramCtrl>,
	public CComCoClass<CTangramCtrl, &CLSID_TangramCtrl>,
	public IDispatchImpl<ITangramCtrl, &IID_ITangramCtrl, &LIBID_Tangram, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	HWND						m_hEclipseViewWnd;
	IWndPage*					m_pPage;
	IWndNode*					m_pCurNode;
	map<CString, CString>		m_mapTangramInfo;
	map<CString, HWND>			m_mapTangramHandle;
	map<CString, IWndFrame*>	m_mapTangramFrame;

	CEclipseWnd*				m_pEclipseWnd;

#pragma warning(push)
#pragma warning(disable: 4355) // 'this' : used in base member initializer list

	CTangramCtrl();

#pragma warning(pop)

DECLARE_OLEMISC_STATUS(OLEMISC_RECOMPOSEONRESIZE |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_INSIDEOUT //|
	//OLEMISC_ACTIVATEWHENVISIBLE |
	//OLEMISC_SETCLIENTSITEFIRST
)

DECLARE_REGISTRY_RESOURCEID(IDR_TANGRAMCTRL)

DECLARE_WND_CLASS(_T("Tangram Control Class"))
BEGIN_COM_MAP(CTangramCtrl)
	COM_INTERFACE_ENTRY(ITangramCtrl)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IViewObject2)
	COM_INTERFACE_ENTRY(IViewObject)
	COM_INTERFACE_ENTRY(IOleInPlaceObject)
	COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
	COM_INTERFACE_ENTRY(IOleObject)
	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
END_COM_MAP()

BEGIN_PROP_MAP(CTangramCtrl)
END_PROP_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct();
	void OnFinalMessage(HWND hWnd);
	void FinalRelease();

// ITangramCtrl
	STDMETHOD(get_HWND)(long* pVal);
	STDMETHOD(GetCtrlText)(long nHandle, BSTR bstrNodeName, BSTR bstrCtrlName, BSTR* bstrVal);
	STDMETHOD(InitCtrl)(BSTR bstrXml);
	STDMETHOD(put_TangramHandle)(BSTR bstrHandleName, LONG newVal);
	STDMETHOD(Extend)(BSTR bstrFrameName, BSTR bstrKey, BSTR bstrXml, IWndNode** ppNode);

private:
	HWND m_hHostWnd;
};

OBJECT_ENTRY_AUTO(__uuidof(TangramCtrl), CTangramCtrl)
