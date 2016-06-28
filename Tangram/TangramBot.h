// TangramBot.h : CTangramBot 

#pragma once
#include "resource.h"       
#include "Tangram.h"
#include "_ITangramBotEvents_CP.h"

using namespace ATL;

// CTangramBot

class ATL_NO_VTABLE CTangramBot :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CTangramBot, &CLSID_TangramBot>,
	public IConnectionPointContainerImpl<CTangramBot>,
	public CProxy_ITangramBotEvents<CTangramBot>,
	public IDispatchImpl<ITangramBot, &IID_ITangramBot, &LIBID_Tangram, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CTangramBot();
	~CTangramBot();


	ITangram*	m_pTangram;
	
DECLARE_REGISTRY_RESOURCEID(IDR_TANGRAMBOT)

DECLARE_NOT_AGGREGATABLE(CTangramBot)

BEGIN_COM_MAP(CTangramBot)
	COM_INTERFACE_ENTRY(ITangramBot)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
END_COM_MAP()

BEGIN_CONNECTION_POINT_MAP(CTangramBot)
	CONNECTION_POINT_ENTRY(__uuidof(_ITangramBotEvents))
END_CONNECTION_POINT_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct();

	void FinalRelease()
	{
	}

public:



};

OBJECT_ENTRY_AUTO(__uuidof(TangramBot), CTangramBot)
