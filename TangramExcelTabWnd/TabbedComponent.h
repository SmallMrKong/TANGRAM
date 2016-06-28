// TabbedComponent.h : Declaration of the CTabbedComponent

#pragma once
#include "resource.h"       // main symbols

#include "TangramExcelTabWnd_i.h"
#include "Tangram.h"
#include "TangramCategory.h"

// CTabbedComponent

class ATL_NO_VTABLE CTabbedComponent :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CTabbedComponent, &CLSID_TabbedComponent>,
	public IDispatchImpl<ICreator, &__uuidof(ICreator)>
{
public:
	CTabbedComponent()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_TABBEDCOMPONENT)

DECLARE_NOT_AGGREGATABLE(CTabbedComponent)

BEGIN_COM_MAP(CTabbedComponent)
	COM_INTERFACE_ENTRY(ICreator)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


BEGIN_CATEGORY_MAP(CTabbedComponent)
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
	STDMETHOD(Create)(long hParentWnd, IWndNode* pNode, long* hWnd);
	STDMETHOD(get_Names)(BSTR* pVal);
	STDMETHOD(get_Tags)(BSTR bstrName, BSTR* pVal);
};

OBJECT_ENTRY_AUTO(__uuidof(TabbedComponent), CTabbedComponent)
