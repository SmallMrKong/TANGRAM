// TangramBot.cpp : CTangramBot µÄÊµÏÖ

#include "stdafx.h"
#include "CloudAddinApp.h"
#include "TangramBot.h"


// CTangramBot

CTangramBot::CTangramBot()
{
	CComPtr<ITangram> pTangram;
	pTangram.CoCreateInstance(L"tangram.tangram.1");
	m_pTangram = pTangram.Detach();
	m_pTangram->AddRef();
}

CTangramBot::~CTangramBot()
{
	ATLTRACE(_T("TangramBot Release:%x\n"), this);
}

HRESULT CTangramBot::FinalConstruct()
{
	//IStream* pStream = 0;
	//HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_IDispatch, theApp.m_pTangram, &pStream);
	//if (hr == S_OK)
	//{
	//	IDispatch* pTarget = nullptr;
	//	hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_IDispatch, (LPVOID *)&pTarget);
	//	if (hr == S_OK&&pTarget)
	//	{
	//		//*ppv = pTarget;
	//		//return S_OK;
	//	}
	//}
	return S_OK;
}
