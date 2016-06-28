// CloudAddinRestObj.h : CRestObj 

#pragma once
#include "resource.h"    
#include "Tangram.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE 平台(如不提供完全 DCOM 支持的 Windows Mobile 平台)上无法正确支持单线程 COM 对象。定义 _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA 可强制 ATL 支持创建单线程 COM 对象实现并允许使用其单线程 COM 对象实现。rgs 文件中的线程模型已被设置为“Free”，原因是该模型是非 DCOM Windows CE 平台支持的唯一线程模型。"
#endif

using namespace ATL;

// CRestObj

class ATL_NO_VTABLE CRestObj :
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDispatchImpl<IRestObject, &IID_IRestObject, &LIBID_Tangram, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CRestObj();
	~CRestObj();

DECLARE_NO_REGISTRY()

BEGIN_COM_MAP(CRestObj)
	COM_INTERFACE_ENTRY(IRestObject)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

	HWND					m_hNotifyWnd;
	IRestNotify*			m_pRestNotify;
	IWndNode*				m_pCloudAddinUINode;
protected:
	ULONG InternalAddRef() { return 1; }
	ULONG InternalRelease() { return 1; }
private:
	int						m_nState;
	CString					m_strKey;
	map<CString, CString>	m_mapHeaders;
public:
	STDMETHOD(UploadFile)(VARIANT_BOOL bUpload, BSTR bstrServerURL, BSTR bstrRequest, BSTR bstrFilePath);
	STDMETHOD(Notify)(long nNotify);
	STDMETHOD(get_CloudAddinRestNotify)(IRestNotify** pVal);
	STDMETHOD(put_CloudAddinRestNotify)(IRestNotify* newVal);
	STDMETHOD(get_WndNode)(IWndNode** pVal);
	STDMETHOD(put_WndNode)(IWndNode* newVal);
	STDMETHOD(get_NotifyHandle)(LONGLONG* pVal);
	STDMETHOD(put_NotifyHandle)(LONGLONG newVal);
	STDMETHOD(get_Header)(BSTR bstrHeaderName, BSTR* pVal);
	STDMETHOD(put_Header)(BSTR bstrHeaderName, BSTR newVal);
	STDMETHOD(ClearHeaders)();
	STDMETHOD(AsyncRest)(int nMethod, BSTR bstrURL, BSTR bstrData, LONGLONG hNotify);
	STDMETHOD(RestData)(int nMethod, BSTR bstrServerURL, BSTR bstrRequest, BSTR bstrData, LONGLONG hNotify);
	STDMETHOD(Clone)(IRestObject* pTargetObj);
	STDMETHOD(get_RestKey)(BSTR* pVal);
	STDMETHOD(put_RestKey)(BSTR newVal);
	STDMETHOD(get_State)(int* pVal);
	STDMETHOD(put_State)(int newVal);
};

