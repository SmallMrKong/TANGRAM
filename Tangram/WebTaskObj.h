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

// CloudAddinGlobalTaskObj.h : CWebTaskObj implementation file

#pragma once
#include "resource.h"      
#include "Tangram.h"
#include <assert.h>
#include <mshtmdid.h>
#include <wininet.h>
#include "JSExtender.h"

const TCHAR c_szDefaultUrl[] = _T("about:blank");

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE 平台(如不提供完全 DCOM 支持的 Windows Mobile 平台)上无法正确支持单线程 COM 对象。定义 _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA 可强制 ATL 支持创建单线程 COM 对象实现并允许使用其单线程 COM 对象实现。rgs 文件中的线程模型已被设置为“Free”，原因是该模型是非 DCOM Windows CE 平台支持的唯一线程模型。"
#endif

// CTaskNotify
class ATL_NO_VTABLE CTaskNotify :
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDispatchImpl<ITaskNotify, &IID_ITaskNotify, &LIBID_Tangram, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CTaskNotify()
	{
	}

	DECLARE_NO_REGISTRY()

	DECLARE_NOT_AGGREGATABLE(CTaskNotify)

	BEGIN_COM_MAP(CTaskNotify)
		COM_INTERFACE_ENTRY(ITaskNotify)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

public:
	STDMETHOD(Notify)(BSTR bstrNotify);
	STDMETHOD(NotifyEx)(VARIANT varNotify);
};

class CTaskMessage
{
public:
	CTaskMessage(void);
	virtual ~CTaskMessage(void);

	int					m_nType;
	CString				m_strXml;
	CString				m_strUrl;
	CString				m_strObjID;

	ITaskNotify*		m_pNotify;
	CTaskNotify*		m_pTaskNotify;

	void Init(CString strXml);
};

// CWebTaskObj

class ATL_NO_VTABLE CWebTaskObj :
	public CJSExtender,
	public IPropertyNotifySink,
	public CComObjectRootBase,
	public IDispatchImpl<IWebTaskObj, &IID_IWebTaskObj, &LIBID_Tangram, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CWebTaskObj();
	~CWebTaskObj();

	HANDLE						m_hHandle;
	DWORD						m_dwThreadID;
	CString						m_strName;
	DECLARE_NOT_AGGREGATABLE(CWebTaskObj)

BEGIN_COM_MAP(CWebTaskObj)
	COM_INTERFACE_ENTRY(IWebTaskObj)
	COM_INTERFACE_ENTRY(IPropertyNotifySink)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	void ProcessMsg(CTaskMessage* pMsg);

	// IPropertyNotifySink methods
	STDMETHOD(OnChanged)(DISPID dispID);
	STDMETHOD(OnRequestEdit)(DISPID dispID);
public:
	STDMETHOD(get_XML)(BSTR* pVal);
	STDMETHOD(put_XML)(BSTR newVal);
	STDMETHOD(get_Extender)(BSTR bstrExtenderName, IDispatch** pVal);
	STDMETHOD(put_Extender)(BSTR bstrExtenderName, IDispatch* newVal);
	STDMETHOD(Run)(void);
	STDMETHOD(Quit)(void);
	STDMETHOD(get_Type)(int* pVal);
	STDMETHOD(put_Type)(int newVal);
	STDMETHOD(InitWebConnection)(BSTR bstrURL, ITaskNotify* pDispNotify);
	STDMETHOD(TangramAction)(BSTR bstrXml, ITaskNotify* pDispNotify);
	STDMETHOD(execScript)(BSTR bstrScript, VARIANT* varRet);

	void Lock() {}
	void Unlock() {}
	CString GetText();

protected:
	ULONG InternalAddRef() { return 1; }
	ULONG InternalRelease() { return 1; }

private:
	bool						m_bIsBlank;
	bool						m_bWebInit;
	DWORD						m_dwCookie;
	HRESULT						m_hrConnected;
	READYSTATE					m_lReadyState;
	INTERNET_SCHEME				m_nScheme;
	CString						m_szURL;
	LPCONNECTIONPOINT			m_pCP;
	ITangram*					m_pTangram;
	IWebTaskObj*				m_pWebTaskObj;
	IHTMLWindow2*				m_pHTMLWindow2;
	IHTMLDocument2*				m_pHtmlDocument2;
	map<CString, IDispatch*>	m_mapExternalObj;
	map<CString, IStream*>		m_mapStreamObj;
	vector<CTaskMessage*>		m_pMsgList;
	HRESULT Init(CString szURL);
	// Persistence helpers
	HRESULT LoadURLFromFile();
	HRESULT LoadURLFromMoniker();
	HRESULT UnLoad();
};

class ATL_NO_VTABLE CTaskObj :
	public IObjectWithSiteImpl<CTaskObj>,
	public CComObjectRootBase,
	public IDispatchImpl<ITaskObj, &IID_ITaskObj, &LIBID_Tangram, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CTaskObj();

	CString						m_strXml;
	map<CString, IDispatch*>	m_mapObjs;

	void Lock() {}
	void Unlock() {}
	DECLARE_NO_REGISTRY()

	BEGIN_COM_MAP(CTaskObj)
		COM_INTERFACE_ENTRY(ITaskObj)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(IObjectWithSite)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

protected:
	ULONG InternalAddRef() { return 1; }
	ULONG InternalRelease() { return 1; }

public:
	STDMETHOD(get_TaskXML)(BSTR* pVal);
	STDMETHOD(put_TaskXML)(BSTR newVal);
	STDMETHOD(Execute)(ITaskNotify* pCallBack, IDispatch* pStateDisp, IDispatch** DispRet);
	STDMETHOD(CreateNode)(TaskNodeType NodeType, BSTR NodeName, BSTR bstrXml);
	STDMETHOD(CreateNode2)(TaskNodeType NodeType, BSTR NodeName, ITaskObj* pTangramTaskObj);
	STDMETHOD(get_TaskParticipantObj)(BSTR bstrID, IDispatch** pVal);
	STDMETHOD(put_TaskParticipantObj)(BSTR bstrID, IDispatch* newVal);
};
