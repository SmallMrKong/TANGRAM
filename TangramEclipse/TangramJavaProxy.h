// TangramJavaProxy.h : CTangramJavaProxy ������

#pragma once
#include "resource.h"       // ������


#include "tangram.h"
#include "TangramEclipse_i.h"
#include "_ITangramJavaProxyEvents_CP.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE ƽ̨(�粻�ṩ��ȫ DCOM ֧�ֵ� Windows Mobile ƽ̨)���޷���ȷ֧�ֵ��߳� COM ���󡣶��� _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA ��ǿ�� ATL ֧�ִ������߳� COM ����ʵ�ֲ�����ʹ���䵥�߳� COM ����ʵ�֡�rgs �ļ��е��߳�ģ���ѱ�����Ϊ��Free����ԭ���Ǹ�ģ���Ƿ� DCOM Windows CE ƽ̨֧�ֵ�Ψһ�߳�ģ�͡�"
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
