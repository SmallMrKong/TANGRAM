// TangramJavaProxy.h : CTangramJavaProxy 的声明

#pragma once
#include "resource.h"       // 主符号


#include "tangram.h"
#include "TangramEclipse_i.h"
#include "_ITangramJavaProxyEvents_CP.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE 平台(如不提供完全 DCOM 支持的 Windows Mobile 平台)上无法正确支持单线程 COM 对象。定义 _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA 可强制 ATL 支持创建单线程 COM 对象实现并允许使用其单线程 COM 对象实现。rgs 文件中的线程模型已被设置为“Free”，原因是该模型是非 DCOM Windows CE 平台支持的唯一线程模型。"
#endif

using namespace ATL;


// CTangramJavaProxy

class ATL_NO_VTABLE CTangramJavaProxy :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CTangramJavaProxy, &CLSID_TangramJavaProxy>,
	public IConnectionPointContainerImpl<CTangramJavaProxy>,
	public CProxy_ITangramJavaProxyEvents<CTangramJavaProxy>,
	public IDispatchImpl<IJavaProxy, &__uuidof(IJavaProxy), &LIBID_Tangram, /* wMajor = */ 1>
{
public:
	CTangramJavaProxy();
	~CTangramJavaProxy();

	DECLARE_REGISTRY_RESOURCEID(IDR_TANGRAMJAVAPROXY)

	DECLARE_NOT_AGGREGATABLE(CTangramJavaProxy)

	BEGIN_COM_MAP(CTangramJavaProxy)
		COM_INTERFACE_ENTRY(IJavaProxy)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(CTangramJavaProxy)
		CONNECTION_POINT_ENTRY(__uuidof(_ITangramJavaProxyEvents))
	END_CONNECTION_POINT_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct();

	void FinalRelease();

	// IJavaProxy Methods
public:
	STDMETHOD(InitEclipse)();
	STDMETHOD(InitComProxy)();
};

OBJECT_ENTRY_AUTO(__uuidof(TangramJavaProxy), CTangramJavaProxy)
