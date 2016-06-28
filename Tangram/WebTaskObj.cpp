// CloudAddinGlobalTaskObj.cpp : CWebTaskObj µÄÊµÏÖ

#include "stdafx.h"
#include "CloudAddinApp.h"
#include "CloudAddinCore.h"
#include "CloudAddinApp.h"
#include "WebTaskObj.h"

// CTaskNotify

STDMETHODIMP CTaskNotify::Notify(BSTR bstrNotify)
{
	return S_OK;
}

STDMETHODIMP CTaskNotify::NotifyEx(VARIANT varNotify)
{
	return S_OK;
}

CTaskMessage::CTaskMessage(void)
{
	m_nType			= 0;
	m_strObjID		= _T("");
	m_strXml		= _T("");
	m_strUrl		= _T("");
	m_pNotify		= nullptr;
	m_pTaskNotify	= nullptr;
}

CTaskMessage::~CTaskMessage(void)
{
}

void CTaskMessage::Init(CString strXml)
{
	CTangramXmlParse m_Parse;
	if (m_Parse.LoadXml(strXml))
	{
		m_strObjID = m_Parse.attr(_T("ObjID"), _T(""));
		if (m_strObjID != _T(""))
		{
		}
	}
}

// CWebTaskObj

CWebTaskObj::CWebTaskObj()
{
	m_hHandle		= 0;
	m_dwThreadID	= 0;
	m_strName		= _T("");
	m_bWebInit		= false;
	m_pTangram		= nullptr;
	m_pWebTaskObj	= nullptr;
}

CWebTaskObj::~CWebTaskObj()
{
	UnLoad();
	auto it = theApp.m_pHostCore->m_mapWebTask.find(m_strName);
	if (it != theApp.m_pHostCore->m_mapWebTask.end())
	{
		theApp.m_pHostCore->m_mapWebTask.erase(it);
	}
	for (auto it : m_mapExternalObj)
	{
		it.second->Release();
	}
	m_mapExternalObj.clear();
}

STDMETHODIMP CWebTaskObj::get_XML(BSTR* pVal)
{
	return S_OK;
}

STDMETHODIMP CWebTaskObj::put_XML(BSTR newVal)
{
	return S_OK;
}

void CWebTaskObj::ProcessMsg(CTaskMessage* pMsg)
{
	::PostThreadMessage(::GetCurrentThreadId(), WM_USER_TANGRAMTASK, 0, 1);
}

STDMETHODIMP CWebTaskObj::OnChanged(DISPID dispID)
{
	HRESULT hr;

	if (DISPID_READYSTATE == dispID)
	{
		// check the value of the readystate property
		assert(m_pHtmlDocument2);

		VARIANT vResult = { 0 };
		EXCEPINFO excepInfo;
		UINT uArgErr;

		DISPPARAMS dp = { NULL, NULL, 0, 0 };
		if (SUCCEEDED(hr = m_pHtmlDocument2->Invoke(DISPID_READYSTATE, IID_NULL, LOCALE_SYSTEM_DEFAULT,
			DISPATCH_PROPERTYGET, &dp, &vResult, &excepInfo, &uArgErr)))
		{
			assert(VT_I4 == V_VT(&vResult));
			m_lReadyState = (READYSTATE)V_I4(&vResult);
			switch (m_lReadyState)
			{
			case READYSTATE_UNINITIALIZED:	//= 0,
				break;
			case READYSTATE_LOADING: //	= 1,
				break;
			case READYSTATE_LOADED:	// = 2,
				break;
			case READYSTATE_INTERACTIVE: //	= 3,
				break;
			case READYSTATE_COMPLETE: // = 4
				PostThreadMessage(GetCurrentThreadId(), WM_TANGRAM_WEBNODEDOCCOMPLETE, (WPARAM)0, (LPARAM)0);
				break;
			}

			VariantClear(&vResult);
		}
	}

	return NOERROR;
}

STDMETHODIMP CWebTaskObj::OnRequestEdit(DISPID dispID)
{
	return NOERROR;
}

STDMETHODIMP CWebTaskObj::Run(void)
{
	structured_task_group tasks;
	IStream* pStream = 0;
	HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_ITangram, theApp.m_pTangram, &pStream);
	IStream* pStream2 = 0;
	hr = ::CoMarshalInterThreadInterfaceInStream(IID_IWebTaskObj, (IWebTaskObj*)this, &pStream2);
	for (auto it : m_mapExternalObj)
	{
		IStream* pStream = 0;
		hr = ::CoMarshalInterThreadInterfaceInStream(IID_IDispatch, it.second, &pStream);
		if (hr == S_OK&&pStream)
		{
			m_mapStreamObj[it.first] = pStream;
		}
	}
	auto task = make_task([&pStream, &pStream2, this]
	{
		CoInitializeEx(NULL, COINIT_MULTITHREADED);
		m_hHandle = ::GetCurrentThread();
		m_dwThreadID = ::GetCurrentThreadId();
		ITangram* pCloudAddin = nullptr;
		HRESULT hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_ITangram, (LPVOID *)&pCloudAddin);
		hr = ::CoGetInterfaceAndReleaseStream(pStream2, IID_IWebTaskObj, (LPVOID *)&m_pWebTaskObj);
		MSG msg;
		while (GetMessage(&msg, NULL, 0, 0))
		{
			if (WM_QUIT == msg.message)
			{
				break;
			}
			if (msg.hwnd == NULL)
			{
				switch (msg.message)
				{
				case WM_USER_TANGRAMTASK:
				{
					switch (msg.lParam)
					{
					case 0:
					{
						CTaskMessage* pMsg = (CTaskMessage*)msg.wParam;
						if (pMsg != NULL)
						{
							CString strURL = pMsg->m_strUrl;
							if (strURL != _T(""))
							{
								m_pTangram = pCloudAddin;
								Init(strURL);
							}
							delete pMsg;
						}
						else
						{
							m_bWebInit = true;
						}
					}
					break;
					case 1:
					case 2:
					{
						CTaskMessage* pTangramGlobalTaskMessage = NULL;
						theApp.Lock();
						if (m_pMsgList.size())
						{
							std::vector<CTaskMessage*>::iterator it = m_pMsgList.begin();
							pTangramGlobalTaskMessage = *it;
							m_pMsgList.erase(it);
						}
						theApp.Unlock();;
						if (pTangramGlobalTaskMessage)
						{
							switch (pTangramGlobalTaskMessage->m_nType)
							{
							case 0:
								ProcessMsg(pTangramGlobalTaskMessage);
								break;
							case 1:
							{
								CComVariant v;
								m_pHTMLWindow2->execScript(pTangramGlobalTaskMessage->m_strXml.AllocSysString(), 0, &v);
								if (pTangramGlobalTaskMessage->m_pNotify)
									pTangramGlobalTaskMessage->m_pNotify->NotifyEx(v);
							}
							break;
							default:
								break;
							}

							delete pTangramGlobalTaskMessage;
						}
					}
					break;
					default:
					{
						CTaskMessage* pMsg = (CTaskMessage*)msg.wParam;
						if (pMsg)
						{
							delete pMsg;
						}
					}
					break;
					}
				}
				break;
				default:
					break;
				}
			}
			else
			{
				DispatchMessage(&msg);
			}
		}
		CoUninitialize();
		if (WM_QUIT == msg.message)
		{
			auto it = theApp.m_pHostCore->m_mapWebTask.find(m_strName);
			if (it != theApp.m_pHostCore->m_mapWebTask.end())
				theApp.m_pHostCore->m_mapWebTask.erase(it);
			delete this;
		}
	});
	tasks.run(task);

	return S_OK;
}

STDMETHODIMP CWebTaskObj::Quit(void)
{
	::PostThreadMessage(::GetCurrentThreadId(), WM_QUIT, 0, 0);

	return S_OK;
}

STDMETHODIMP CWebTaskObj::get_Type(int* pVal)
{
	// TODO: Add your implementation code here

	return S_OK;
}

STDMETHODIMP CWebTaskObj::put_Type(int newVal)
{
	// TODO: Add your implementation code here

	return S_OK;
}

STDMETHODIMP CWebTaskObj::InitWebConnection(BSTR bstrURL, ITaskNotify* pDispNotify)
{
	if (m_dwThreadID != 0)
	{
		CTaskMessage* pMsg = new CTaskMessage();
		pMsg->m_strUrl = OLE2T(bstrURL);
		if (pDispNotify)
		{
			pMsg->m_pNotify = pDispNotify;
			pMsg->m_pNotify->AddRef();
		}
		::PostThreadMessage(m_dwThreadID, WM_USER_TANGRAMTASK, (WPARAM)pMsg, 0);
	}
	return S_OK;
}

STDMETHODIMP CWebTaskObj::TangramAction(BSTR bstrXml, ITaskNotify* pDispNotify)
{
	CString strXml = OLE2T(bstrXml);
	if (strXml == _T(""))
		return S_OK;
	if (m_dwThreadID != 0)
	{
		Lock();
		CTaskMessage* pMsg = new CTaskMessage();
		pMsg->m_strXml = strXml;
		pMsg->m_pNotify = pDispNotify;
		pMsg->m_pNotify->AddRef();
		m_pMsgList.push_back(pMsg);
		Unlock();
		::PostThreadMessage(m_dwThreadID, WM_USER_TANGRAMTASK, 1, 0);
	}

	return S_OK;
}

STDMETHODIMP CWebTaskObj::execScript(BSTR bstrScript, VARIANT* varRet)
{
	CString strXml = OLE2T(bstrScript);
	if (strXml == _T(""))
		return S_OK;
	if (m_dwThreadID != 0)
	{
		Lock();
		CTaskMessage* pMsg = new CTaskMessage();
		pMsg->m_nType = 1;
		pMsg->m_strXml = strXml;
		//pMsg->m_pNotify = pDispNotify;
		//pMsg->m_pNotify->AddRef();
		m_pMsgList.push_back(pMsg);
		Unlock();
		::PostThreadMessage(m_dwThreadID, WM_USER_TANGRAMTASK, 1, 0);
	}

	return S_OK;
}

STDMETHODIMP CWebTaskObj::get_Extender(BSTR bstrExtenderName, IDispatch** pVal)
{
	theApp.Lock();
	CString strName = OLE2T(bstrExtenderName);
	if (strName == _T(""))
		return S_OK;
	auto it = m_mapExternalObj.find(strName);
	if (it != m_mapExternalObj.end())
	{
		*pVal = it->second;
		(*pVal)->AddRef();
	}
	theApp.Unlock();
	return S_OK;
}

STDMETHODIMP CWebTaskObj::put_Extender(BSTR bstrExtenderName, IDispatch* newVal)
{
	theApp.Lock();
	CString strName = OLE2T(bstrExtenderName);
	if (strName == _T(""))
		return S_OK;
	auto it = m_mapExternalObj.find(strName);
	if (it == m_mapExternalObj.end())
	{
		m_mapExternalObj[strName] = newVal;
		newVal->AddRef();
	}
	theApp.Unlock();
	return S_OK;
}

// CTaskObj
CTaskObj::CTaskObj()
{
	m_strXml = _T("");
}

STDMETHODIMP CTaskObj::get_TaskXML(BSTR* pVal)
{
	// TODO: Add your implementation code here
	*pVal = CComBSTR(m_strXml);
	return S_OK;
}


STDMETHODIMP CTaskObj::put_TaskXML(BSTR newVal)
{
	// TODO: Add your implementation code here
	m_strXml = OLE2T(newVal);
	return S_OK;
}

STDMETHODIMP CTaskObj::Execute(ITaskNotify* pCallBack, IDispatch* pStateDisp, IDispatch** DispRet)
{
	if (m_strXml == _T(""))
		return S_FALSE;
	int nCount = 0;
	CTangramXmlParse m_Parse;
	if (m_Parse.LoadXml(m_strXml))
	{
		CTangramXmlParse* pTaskChild = m_Parse.GetChild(_T("TangramTask"));
		if (pTaskChild == NULL)
			pTaskChild = &m_Parse;
		if (pTaskChild)
		{
			CTangramXmlParse* pChildNode = pTaskChild->GetChild(_T("TaskNoWait"));
			if (pChildNode)
			{
				nCount = pChildNode->GetCount();
				if (nCount)
				{
					for (int i = 0; i < nCount; i++)
					{
						structured_task_group tasks;
						IStream* pStream = 0;
						HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pCallBack, &pStream);
						IStream* pStream2 = 0;
						hr = ::CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pStateDisp, &pStream2);
						auto task = make_task([pChildNode, i, pCallBack, pStream, pStream2, pStateDisp, this] {
							// Initialize the COM library on the current thread.
							CoInitializeEx(NULL, COINIT_MULTITHREADED);
							CTangramXmlParse* pChildNode2 = pChildNode->GetChild(i);
							if (pChildNode2)
							{
								CString strXml = pChildNode2->xml();
								int nCount2 = pChildNode2->GetCount();
								if (nCount2)
								{
									ITaskNotify* pEventTarget = nullptr;
									HRESULT hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_ITaskNotify, (LPVOID *)&pEventTarget);
									IDispatch* pEventTarget2 = nullptr;
									hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_IDispatch, (LPVOID *)&pEventTarget2);
									ITask* pPrevTangramTask = nullptr;
									for (int i = 0; i<nCount2; i++)
									{
										CString strObjID = pChildNode2->GetChild(i)->attr(_T("ObjID"), _T(""));
										if (strObjID != _T(""))
										{
											std::map<CString, IDispatch*>::iterator it = m_mapObjs.find(strObjID);
											if (it != m_mapObjs.end())
											{
												CComQIPtr<ITask> pTangramTask(it->second);
												if (pTangramTask != nullptr)
												{
													pTangramTask->Execute(strObjID.AllocSysString(), pEventTarget, pEventTarget2, pPrevTangramTask);
													if (pPrevTangramTask != nullptr)
														pPrevTangramTask->Release();
													if (i<nCount2 - 1)
													{
														pPrevTangramTask = pTangramTask.p;
														pPrevTangramTask->AddRef();
													}
												}
											}
											else
											{
												CComPtr<IDispatch> pDisp;
												hr = pDisp.CoCreateInstance(CComBSTR(strObjID));
												if (hr == S_OK)
												{
													CComQIPtr<ITask> pTangramTask(pDisp.p);
													if (pTangramTask != nullptr)
													{
														pTangramTask->Execute(strObjID.AllocSysString(), pEventTarget, pEventTarget2, pPrevTangramTask);
														if (pPrevTangramTask != nullptr)
															pPrevTangramTask->Release();
														if (i<nCount2 - 1)
														{
															pPrevTangramTask = pTangramTask.p;
															pPrevTangramTask->AddRef();
														}
													}
												}
											}
										}
									}//
								}
							}
							CoUninitialize();
						});
						tasks.run(task);
					}
				}
			}

			pChildNode = pTaskChild->GetChild(_T("TaskNeedWait"));
			if (pChildNode)
			{
				nCount = pChildNode->GetCount();
				if (nCount)
				{
					std::vector<Concurrency::task<void>> mapTasks;

					for (int i = 0; i < nCount; i++)
					{
						IStream* pStream = 0;
						HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pCallBack, &pStream);
						IStream* pStream2 = 0;
						hr = ::CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pStateDisp, &pStream2);
						auto task = Concurrency::create_task([pChildNode, i, pCallBack, pStream, pStream2, pStateDisp, this] {
							CoInitializeEx(NULL, COINIT_MULTITHREADED);
							CTangramXmlParse* pChildNode2 = pChildNode->GetChild(i);
							if (pChildNode2)
							{
								CString strXml = pChildNode2->xml();
								int nCount2 = pChildNode2->GetCount();
								if (nCount2)
								{
									ITaskNotify* pEventTarget = nullptr;
									HRESULT hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_ITaskNotify, (LPVOID *)&pEventTarget);
									IDispatch* pEventTarget2 = nullptr;
									hr = ::CoGetInterfaceAndReleaseStream(pStream2, IID_IDispatch, (LPVOID *)&pEventTarget2);
									ITask* pPrevTangramTask = nullptr;
									for (int i = 0; i<nCount2; i++)
									{
										CString strObjID = pChildNode2->GetChild(i)->attr(_T("ObjID"), _T(""));
										if (strObjID != _T(""))
										{
											std::map<CString, IDispatch*>::iterator it = m_mapObjs.find(strObjID);
											if (it != m_mapObjs.end())
											{
												CComQIPtr<ITask> pTangramTask(it->second);
												if (pTangramTask != nullptr)
												{
													pTangramTask->Execute(strObjID.AllocSysString(), pEventTarget, pEventTarget2, pPrevTangramTask);
													if (pPrevTangramTask != nullptr)
														pPrevTangramTask->Release();
													if (i<nCount2 - 1)
													{
														pPrevTangramTask = pTangramTask.p;
														pPrevTangramTask->AddRef();
													}
												}
											}
											else
											{
												CComPtr<IDispatch> pDisp;
												hr = pDisp.CoCreateInstance(CComBSTR(strObjID));
												if (hr == S_OK)
												{
													CComQIPtr<ITask> pTangramTask(pDisp.p);
													if (pTangramTask != nullptr)
													{
														pTangramTask->Execute(strObjID.AllocSysString(), pEventTarget, pEventTarget2, pPrevTangramTask);
														if (pPrevTangramTask != nullptr)
															pPrevTangramTask->Release();
														if (i<nCount2 - 1)
														{
															pPrevTangramTask = pTangramTask.p;
															pPrevTangramTask->AddRef();
														}
													}
												}
											}
										}
									}
								}
							}
							CoUninitialize();
						});
						mapTasks.push_back(task);
					}
					auto joinTask = Concurrency::when_all(mapTasks.begin(), mapTasks.end());
					joinTask.wait();
				}
			}

			pChildNode = pTaskChild->GetChild(_T("Default"));
			if (pChildNode)
			{
				nCount = pChildNode->GetCount();
				if (nCount)
				{
					for (int i = 0; i < nCount; i++)
					{
						structured_task_group tasks;
						IStream* pStream = 0;
						HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pCallBack, &pStream);
						IStream* pStream2 = 0;
						hr = ::CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pStateDisp, &pStream2);
						auto task = make_task([pChildNode, i, pCallBack, pStream, pStream2, pStateDisp, this] 
						{
							// Initialize the COM library on the current thread.
							CoInitializeEx(NULL, COINIT_MULTITHREADED);
							CTangramXmlParse* pChildNode2 = pChildNode->GetChild(i);
							if (pChildNode2)
							{
								CString strXml = pChildNode2->xml();
								int nCount2 = pChildNode2->GetCount();
								if (nCount2)
								{
									ITaskNotify* pEventTarget = nullptr;
									HRESULT hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_ITaskNotify, (LPVOID *)&pEventTarget);
									IDispatch* pEventTarget2 = nullptr;
									hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_IDispatch, (LPVOID *)&pEventTarget2);
									ITask* pPrevTangramTask = nullptr;
									for (int i = 0; i<nCount2; i++)
									{
										CString strObjID = pChildNode2->GetChild(i)->attr(_T("ObjID"), _T(""));
										if (strObjID != _T(""))
										{
											std::map<CString, IDispatch*>::iterator it = m_mapObjs.find(strObjID);
											if (it != m_mapObjs.end())
											{
												CComQIPtr<ITask> pTangramTask(it->second);
												if (pTangramTask != nullptr)
												{
													pTangramTask->Execute(strObjID.AllocSysString(), pEventTarget, pEventTarget2, pPrevTangramTask);
													if (pPrevTangramTask != nullptr)
														pPrevTangramTask->Release();
													if (i<nCount2 - 1)
													{
														pPrevTangramTask = pTangramTask.p;
														pPrevTangramTask->AddRef();
													}
												}
											}
											else
											{
												CComPtr<IDispatch> pDisp;
												hr = pDisp.CoCreateInstance(CComBSTR(strObjID));
												if (hr == S_OK)
												{
													CComQIPtr<ITask> pTangramTask(pDisp.p);
													if (pTangramTask != nullptr)
													{
														pTangramTask->Execute(strObjID.AllocSysString(), pEventTarget, pEventTarget2, pPrevTangramTask);
														if (pPrevTangramTask != nullptr)
															pPrevTangramTask->Release();
														if (i<nCount2 - 1)
														{
															pPrevTangramTask = pTangramTask.p;
															pPrevTangramTask->AddRef();
														}
													}
												}
											}
										}
									}
								}
							}
							CoUninitialize();
						});
						tasks.run_and_wait(task);
					}
				}
			}
		}
		else
		{
		}
	}

	return S_OK;
}

STDMETHODIMP CTaskObj::CreateNode2(TaskNodeType nType, BSTR NodeName, ITaskObj* pTangramTaskObj)
{
	CString strName = OLE2T(NodeName);
	if (pTangramTaskObj && nType != TaskNodeType::TANGRAM && nType != TaskNodeType::UCMA && strName != _T(""))
	{
		CComBSTR strXml(L"");
		HRESULT hr = pTangramTaskObj->get_TaskXML(&strXml);
		CString _strXml = OLE2T(strXml);
		if (_strXml != _T(""))
		{
			return CreateNode(nType, NodeName, strXml);
		}
	}
	return S_OK;
}

STDMETHODIMP CTaskObj::CreateNode(TaskNodeType nType, BSTR NodeName, BSTR bstrXml)
{
	CString strNodeName = OLE2T(NodeName);
	if (strNodeName == _T(""))
		return S_OK;

	if (nType == TaskNodeType::TANGRAM || nType == TaskNodeType::UCMA)
	{
		CString m_strXMLFormat = _T("<Tangram-%s><TangramTask></TangramTask></Tangram-%s>");
		if (nType == TaskNodeType::UCMA)
			m_strXMLFormat = L"<Ucma-%s><TangramTask></TangramTask></Ucma-%s>";
		m_strXml.Format(m_strXMLFormat, strNodeName, strNodeName);
		return S_OK;
	}
	CString strXml = OLE2T(bstrXml);
	if (strXml == _T("") || strXml.Find(_T(" ObjID")) == -1)
		return S_OK;
	if (m_strXml == _T(""))
		m_strXml = _T("<TangramDefault><TangramTask></TangramTask></TangramDefault>");

	CString strGUID = theApp.GetNewGUID();

	CString strXml2 = _T("");
	strXml2.Format(_T("<%s>%s</%s>"), strNodeName, strXml, strNodeName);

	CTangramXmlParse m_Parse;
	CTangramXmlParse* pTaskNode = nullptr;
	CTangramXmlParse* pTaskNode2 = nullptr;
	if (m_Parse.LoadXml(strXml))
	{
		CString _strXml2 = _T("");
		CString strKey = _T("");
		pTaskNode = m_Parse.GetChild(_T("TangramTask"));
		if (pTaskNode == nullptr)
			return S_OK;
		switch (nType)
		{
		case TaskNodeType::TaskNoWait:
		{
			pTaskNode2 = pTaskNode->GetChild(_T("TaskNoWait"));
			if (pTaskNode2 == nullptr)
			{
				pTaskNode->AddNode(_T("TaskNoWait"))->put_text(strGUID);
				strKey = strXml2;
			}
		}
		break;
		case TaskNodeType::TaskNeedWait:
		{
			pTaskNode2 = pTaskNode->GetChild(_T("TaskNeedWait"));
			if (pTaskNode2 == nullptr)
			{
				pTaskNode->AddNode(_T("TaskNeedWait"))->put_text(strGUID);
				strKey = strXml2;
			}
		}
		break;
		case TaskNodeType::DefaultNode:
		{
			pTaskNode2 = pTaskNode->GetChild(_T("Default"));
			if (pTaskNode2 == nullptr)
			{
				pTaskNode->AddNode(_T("Default"))->put_text(strGUID);
				strKey = strXml2;
			}
		}
		break;
		}
		if (pTaskNode2)
		{
			CTangramXmlParse* pNode = pTaskNode2->GetChild(strNodeName);
			if (pNode == nullptr)
			{
				pTaskNode2->AddNode(strNodeName)->put_text(strGUID);
				strKey = strXml;
			}
			else
			{
				pNode->AddNode(strGUID);
				_strXml2 = m_Parse.xml();
				CString s = _T("");
				s.Format(_T("<%s/>"), strGUID);
				if (_strXml2.Find(s) != -1)
				{
					_strXml2.Replace(s, strXml);
				}
				else
				{
					s.Format(_T("<%s />"), strGUID);
					if (_strXml2.Find(s) != -1)
					{
						_strXml2.Replace(s, strXml);
					}
					else
					{
						s.Format(_T("<%s></%s>"), strGUID, strGUID);
						if (_strXml2.Find(s) != -1)
						{
							_strXml2.Replace(s, strXml);
						}
					}
				}
				m_strXml = _strXml2;
				return S_OK;
			}
		}
		if (strKey != _T(""))
		{
			_strXml2 = m_Parse.xml();
			_strXml2.Replace(strGUID, strKey);
			m_strXml = _strXml2;
		}
	}

	return S_OK;
}

STDMETHODIMP CTaskObj::get_TaskParticipantObj(BSTR bstrID, IDispatch** pVal)
{
	CString strID = OLE2T(bstrID);
	if (strID != _T(""))
	{
		std::map<CString, IDispatch*>::iterator it = this->m_mapObjs.find(strID);
		if (it != m_mapObjs.end())
		{
			*pVal = it->second;
			(*pVal)->AddRef();
		}
	}

	return S_OK;
}

STDMETHODIMP CTaskObj::put_TaskParticipantObj(BSTR bstrID, IDispatch* newVal)
{
	CString strID = OLE2T(bstrID);
	if (strID != _T(""))
	{
		std::map<CString, IDispatch*>::iterator it = this->m_mapObjs.find(strID);
		if (it != m_mapObjs.end())
		{
			m_mapObjs.erase(it);
		}
		m_mapObjs[strID] = newVal;
	}

	return S_OK;
}

// Initialize the app. Load MSHTML, hook up property notification sink, etc
HRESULT CWebTaskObj::Init(CString szURL)
{
	HRESULT hr;
	LPCONNECTIONPOINTCONTAINER pCPC = nullptr;
	CComPtr<IDispatch> pScriptDisp;

	if (szURL == _T(""))
	{
		m_nScheme = INTERNET_SCHEME_HTTP;
	}
	else
	{
		URL_COMPONENTS urlComponents;
		DWORD dwFlags = 0;
		m_nScheme = INTERNET_SCHEME_UNKNOWN;
		ZeroMemory((PVOID)&urlComponents, sizeof(URL_COMPONENTS));
		urlComponents.dwStructSize = sizeof(URL_COMPONENTS);

		if (szURL)
		{
			if (InternetCrackUrl(szURL, 0, dwFlags, &urlComponents))
			{
				m_nScheme = urlComponents.nScheme;
			}
		}
	}

	if (FAILED(hr = CoCreateInstance(CLSID_HTMLDocument, NULL, CLSCTX_INPROC_SERVER, IID_IHTMLDocument2, (LPVOID*)&m_pHtmlDocument2)))
	{
		goto Error;
	}

	if (SUCCEEDED(hr = m_pHtmlDocument2->get_parentWindow(&m_pHTMLWindow2)))
	{
	}

	// Hook up sink to catch ready state property change
	if (FAILED(hr = m_pHtmlDocument2->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&pCPC)))
	{
		goto Error;
	}

	if (FAILED(hr = pCPC->FindConnectionPoint(IID_IPropertyNotifySink, &m_pCP)))
	{
		goto Error;
	}

	m_hrConnected = m_pCP->Advise((LPUNKNOWN)(IPropertyNotifySink*)this, &m_dwCookie);

	if (m_pHtmlDocument2->get_Script(&pScriptDisp) == S_OK)
	{
		CComPtr<IDispatchEx> _pScriptEx;
		hr = pScriptDisp->QueryInterface<IDispatchEx>(&_pScriptEx);
		if (hr == S_OK)
		{
			Lock();
			ConnectJSEngine(_pScriptEx);
			CJSExtender::AddObject(_T("Tangram"), m_pTangram);
			if(m_pWebTaskObj)
			CJSExtender::AddObject(_T("WebTask"), m_pWebTaskObj);
			for (auto it : m_mapStreamObj)
			{
				IDispatch* pDisp = nullptr;
				HRESULT hr = ::CoGetInterfaceAndReleaseStream(it.second, IID_IDispatch, (LPVOID *)&pDisp);
				if (hr == S_OK&&pDisp)
				{
					CJSExtender::AddObject(it.first, pDisp);
					CString strEvent = it.first;
					strEvent += _T("_On");
					CJSExtender::SinkObject(strEvent, pDisp);
				}
			}
			Unlock();
		}
	}

	switch (m_nScheme)
	{
	case INTERNET_SCHEME_HTTP:
	case INTERNET_SCHEME_FTP:
	case INTERNET_SCHEME_RES:
	case INTERNET_SCHEME_GOPHER:
	case INTERNET_SCHEME_HTTPS:
	case INTERNET_SCHEME_FILE:
		// load URL using IPersistMoniker
		hr = LoadURLFromMoniker();
		break;
	case INTERNET_SCHEME_NEWS:
	case INTERNET_SCHEME_MAILTO:
	case INTERNET_SCHEME_SOCKS:
		// we don't handle these
		return E_FAIL;
		break;
	default:
		if (m_bIsBlank)
		{
			hr = LoadURLFromMoniker();
			break;
		}
		hr = LoadURLFromFile();
		break;
	}

	return S_OK;
Error:
	if (pCPC)
		pCPC->Release();

	return hr;
}

HRESULT CWebTaskObj::UnLoad()
{
	HRESULT hr = NOERROR;

	// Disconnect from property change notifications
	if (m_pCP)
	{
		if (m_dwCookie)
		{
			hr = m_pCP->Unadvise(m_dwCookie);
			m_dwCookie = 0;
		}

		// Release the connection point
		m_pCP->Release();
		m_pCP = nullptr;
	}

	if (m_pHTMLWindow2)
	{
		DWORD dw = m_pHTMLWindow2->Release();
		while (dw>2)
			dw = m_pHTMLWindow2->Release();
		//CCommonFunction::ClearObject<IUnknown>(m_pHTMLWindow2);
		m_pHTMLWindow2 = nullptr;
	}

	return NOERROR;
}

// Use an asynchronous Moniker to load the specified resource
HRESULT CWebTaskObj::LoadURLFromMoniker()
{
	HRESULT hr;
	// Ask the system for a URL Moniker
	LPMONIKER pMk = nullptr;
	LPBINDCTX pBCtx = nullptr;
	LPPERSISTMONIKER pPMk = nullptr;

	if (FAILED(hr = CreateURLMoniker(NULL, CComBSTR(m_szURL), &pMk)))
	{
		return hr;
	}

	if (FAILED(hr = CreateBindCtx(0, &pBCtx)))
	{
		goto Error;
	}

	// Use MSHTML moniker services to load the specified document
	if (SUCCEEDED(hr = m_pHtmlDocument2->QueryInterface(IID_IPersistMoniker,
		(LPVOID*)&pPMk)))
	{
		hr = pPMk->Load(false, pMk, pBCtx, STGM_READ);
	}

Error:
	if (pMk) pMk->Release();
	if (pBCtx) pBCtx->Release();
	if (pPMk) pPMk->Release();
	return hr;
}

// A more traditional form of persistence.
// MSHTML performs this asynchronously as well.
HRESULT CWebTaskObj::LoadURLFromFile()
{
	HRESULT hr;

	LPPERSISTFILE  pPF;
	// MSHTML supports file persistence for ordinary files.
	if (SUCCEEDED(hr = m_pHtmlDocument2->QueryInterface(IID_IPersistFile, (LPVOID*)&pPF)))
	{
		hr = pPF->Load(CComBSTR(m_szURL), 0);
		pPF->Release();
	}

	return hr;
}

CString CWebTaskObj::GetText()
{
	HRESULT hr;
	assert(m_pHtmlDocument2);
	if (!m_pHtmlDocument2)
	{
		return _T("");
	}

	if (READYSTATE_COMPLETE != m_lReadyState)
	{
		DebugBreak();
		return _T("");
	}

	CComQIPtr<IHTMLDocument3> m_pHTMLDoc(m_pHtmlDocument2);
	IHTMLElement* pDocElem = nullptr;
	m_pHTMLDoc->get_documentElement(&pDocElem);
	CComBSTR bstrTxt(L"");
	hr = pDocElem->get_innerText(&bstrTxt);
	pDocElem->Release();
	return OLE2T(bstrTxt);
}