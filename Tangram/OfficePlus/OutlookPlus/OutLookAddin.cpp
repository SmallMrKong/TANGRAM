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

#include "../../stdafx.h"
#include "../../CloudAddinApp.h"
#include "../../WndFrame.h"
#include "../../WndNode.h"
#include "Tangram.h"
#include "fm20.h"
#include "outlctl.h"
#include "OutLookAddin.h"
#include "OutLookplusEvents.cpp"
#include "OutLookCustomizeFormDlg.h"
#define SECURITY_WIN32
#include <Security.h>
#pragma comment(lib, "Secur32.lib") 
#include <winInet.h>

namespace OfficeCloudPlus
{
	namespace OutLookPlus
	{
		COutLookCloudAddin::COutLookCloudAddin()
		{
			m_strCurrentKey				= _T("");
			m_strDesignExplorerEntryID	= _T("");
			m_strDesignExplorerStoreID	= _T("");
			m_nGetClassDispID = -1;
			m_pTangramInspectorItems = nullptr;
		}

		void COutLookCloudAddin::OnItemLoad(IDispatch* pDisp)
		{
			CInspectorItem* pTangramInspectorItem = new CInspectorItem();
			m_mapTangramInspectorItem[(long)pTangramInspectorItem] = pTangramInspectorItem;
			pTangramInspectorItem->m_pItem = pDisp;
			HRESULT hr = ((COutLookItemEvents*)pTangramInspectorItem)->DispEventAdvise(pDisp);
			if (hr == S_OK)
			{
				pDisp->AddRef();
				//hr = ((COutLookItemEvents_10*)pTangramInspectorItem)->DispEventAdvise(pDisp);
				::PostMessage(m_hHostWnd, WM_TANGRAMITEMLOAD, (WPARAM)pTangramInspectorItem, 0);
			}
			else
			{
				delete pTangramInspectorItem;
			}
		}

		void COutLookCloudAddin::AddTangramPropertyToInspector(_Inspector* pInspector, CString strPropertyName, CString strInfoXml)
		{
			CComPtr<IDispatch> pItem;
			pInspector->get_CurrentItem(&pItem);
			if (pItem)
			{
				BSTR szMember = L"ItemProperties";
				DISPID dispid = -1;
				HRESULT hr = pItem->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_USER_DEFAULT, &dispid);
				if (hr == S_OK)
				{
					DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
					VARIANT result = { 0 };
					EXCEPINFO excepInfo;
					memset(&excepInfo, 0, sizeof excepInfo);
					UINT nArgErr = (UINT)-1;
					HRESULT hr = pItem->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &result, &excepInfo, &nArgErr);
					if (S_OK == hr && VT_DISPATCH == result.vt)
					{
						CComQIPtr<ItemProperties> pProperties(result.pdispVal);
						if(pProperties)
						{
							CComPtr<ItemProperty> pProperty;
							hr = pProperties->Item(CComVariant(strPropertyName), &pProperty);
							if (pProperty == nullptr)
							{
								hr = pProperties->Add(CComBSTR(strPropertyName), OlUserPropertyType::olText, CComVariant(NULL), CComVariant(NULL), &pProperty);
							}
							if(pProperty)
							{
								CComVariant var(strInfoXml);
								pProperty->put_Value(var);
								::VariantClear(&var);
							}
						}
					}
					::VariantClear(&result);
				}
			}
		}

		CString COutLookCloudAddin::GetTangramPropertyFromInspector(_Inspector* pInspector, CString strPropertyName)
		{
			CComPtr<IDispatch> pItem;
			pInspector->get_CurrentItem(&pItem);
			if (pItem)
			{
				return GetTangramPropertyFromItem(pItem, strPropertyName);
			}
			return _T("");
		}

		CString COutLookCloudAddin::GetTangramPropertyFromItem(IDispatch* pItem, CString strPropertyName)
		{
			CString strRet = _T("");
			if (pItem)
			{
				CComQIPtr<IDispatch> pDisp(pItem);
				if (pDisp == NULL)
					return _T("");
				BSTR szMember = L"ItemProperties";
				DISPID dispid = -1;
				HRESULT hr = pItem->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_USER_DEFAULT, &dispid);
				if (hr == S_OK)
				{
					DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
					VARIANT result = { 0 };
					EXCEPINFO excepInfo;
					memset(&excepInfo, 0, sizeof excepInfo);
					UINT nArgErr = (UINT)-1;
					HRESULT hr = pItem->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &result, &excepInfo, &nArgErr);
					if (S_OK == hr && VT_DISPATCH == result.vt)
					{
						CComQIPtr<ItemProperties> pProperties(result.pdispVal);
						if(pProperties)
						{
							CComPtr<ItemProperty> pProperty;
							hr = pProperties->Item(CComVariant(strPropertyName), &pProperty);
							if (hr == S_OK&&pProperty)
							{
								CComVariant var;
								pProperty->get_Value(&var);
								if (var.vt == VT_BSTR)
									strRet = OLE2T(var.bstrVal);
								::VariantClear(&var);
							}
						}
					}
					::VariantClear(&result);
				}
			}
			return strRet;
		}

		_Inspector* COutLookCloudAddin::GetInspector(IDispatch* pDisp)
		{
			if (pDisp)
			{
				BSTR szMember = L"GetInspector";
				DISPID dispid = -1;
				HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_USER_DEFAULT, &dispid);
				if (hr == S_OK)
				{
					DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
					VARIANT result = { 0 };
					EXCEPINFO excepInfo;
					memset(&excepInfo, 0, sizeof excepInfo);
					UINT nArgErr = (UINT)-1;
					HRESULT hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &result, &excepInfo, &nArgErr);
					if (S_OK == hr && VT_DISPATCH == result.vt)
					{
						CComQIPtr<_Inspector> pInspector(result.pdispVal);
						_Inspector* _pInspector = pInspector.p;
						_pInspector->AddRef();
						return _pInspector;
					}
				}
			}
			return nullptr;
		}

		CString COutLookCloudAddin::GetPropertyFromItem(IDispatch* pItem, CString strPropertyName)
		{
			CString strRet = _T("");
			if (pItem)
			{
				BSTR szMember = strPropertyName.AllocSysString();
				DISPID dispid = -1;
				HRESULT hr = pItem->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_USER_DEFAULT, &dispid);
				if (hr == S_OK)
				{
					DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
					VARIANT result = { 0 };
					EXCEPINFO excepInfo;
					memset(&excepInfo, 0, sizeof excepInfo);
					UINT nArgErr = (UINT)-1;
					HRESULT hr = pItem->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &result, &excepInfo, &nArgErr);
					if (S_OK == hr && VT_BSTR == result.vt)
					{
						strRet = OLE2T(result.bstrVal);
					}
					::VariantClear(&result);
				}
			}
			return strRet;
		}

		CString COutLookCloudAddin::GetPropertyFromInspector(_Inspector* pInspector, CString strPropertyName)
		{
			CComPtr<IDispatch> pItem;
			pInspector->get_CurrentItem(&pItem);
			if (pItem)
			{
				return GetPropertyFromItem(pItem, strPropertyName);
			}
			return _T("");
		}

		void COutLookCloudAddin::WriteFolderPropertyToStore(MAPIFolder* _pFolder,CString strSubject, CString strPropertyName, CString strData)
		{
			structured_task_group tasks;
			IStream* pStream = 0;
			HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_MAPIFolder, _pFolder, &pStream);
			auto task = make_task([&pStream, &strSubject, &strPropertyName, &strData]()
			{
				CoInitializeEx(NULL, COINIT_MULTITHREADED);
				if (pStream)
				{
					MAPIFolder* pFolder = nullptr;
					HRESULT hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_MAPIFolder, (LPVOID *)&pFolder);
					if (hr == S_OK&&pFolder)
					{
						CComPtr<OutLook::_StorageItem> pStoreItem;
						hr = pFolder->GetStorage(CComBSTR(strSubject), olIdentifyBySubject, &pStoreItem);
						if (pStoreItem)
						{
							long nSize = 0;
							pStoreItem->get_Size(&nSize);
							CComPtr<OutLook::UserProperties> pUserProperties;
							pStoreItem->get_UserProperties(&pUserProperties);
							CComPtr<OutLook::UserProperty> pUserProperty;
							if (nSize == 0)
							{
								pUserProperties->Add(CComBSTR(strPropertyName), olText, CComVariant(false), CComVariant((int)OlFormatText::olFormatTextText), &pUserProperty);
							}
							else
							{
								pUserProperties->Item(CComVariant(strPropertyName), &pUserProperty);
							}
							if (pUserProperty)
							{
								pUserProperty->put_Value(CComVariant(strData));
								pStoreItem->Save();
							}
						}
					}
				}
				CoUninitialize();
			});
			tasks.run(task);
			//tasks.run_and_wait(task);

			//if (pFolder)
			//{
			//	CComPtr<OutLook::_StorageItem> pStoreItem;
			//	pFolder->GetStorage(CComBSTR(strSubject), olIdentifyBySubject, &pStoreItem);
			//	if (pStoreItem)
			//	{
			//		long nSize = 0;
			//		pStoreItem->get_Size(&nSize);
			//		CComPtr<OutLook::UserProperties> pUserProperties;
			//		pStoreItem->get_UserProperties(&pUserProperties);
			//		CComPtr<OutLook::UserProperty> pUserProperty;
			//		if (nSize == 0)
			//		{
			//			pUserProperties->Add(CComBSTR(strPropertyName), olText, CComVariant(false), CComVariant((int)OlFormatText::olFormatTextText), &pUserProperty);
			//		}
			//		else
			//		{
			//			pUserProperties->Item(CComVariant(strPropertyName), &pUserProperty);
			//		}
			//		if (pUserProperty)
			//		{
			//			pUserProperty->put_Value(CComVariant(strData));
			//			pStoreItem->Save();
			//		}
			//	}
			//}
		}

		CString COutLookCloudAddin::GetFolderPropertyFromStore(MAPIFolder* _pFolder,CString strSubject, CString strPropertyName)
		{
			CString strRet = _T("");
			structured_task_group tasks;
			IStream* pStream = 0;
			HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_MAPIFolder, _pFolder, &pStream);
			auto task = make_task([&pStream, &strSubject, &strPropertyName, &strRet]()
			{
				//CoInitializeEx(NULL, COINIT_MULTITHREADED);
				if (pStream)
				{
					MAPIFolder* pFolder = nullptr;
					HRESULT hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_MAPIFolder, (LPVOID *)&pFolder);
					if (hr == S_OK&&pFolder)
					{
						CComPtr<OutLook::_StorageItem> pStoreItem;
						hr = pFolder->GetStorage(CComBSTR(strSubject), olIdentifyBySubject, &pStoreItem);
						if (pStoreItem)
						{
							long nSize = 0;
							pStoreItem->get_Size(&nSize);
							CComPtr<OutLook::UserProperties> pUserProperties;
							pStoreItem->get_UserProperties(&pUserProperties);
							CComPtr<OutLook::UserProperty> pUserProperty;
							if (nSize)
							{
								pUserProperties->Item(CComVariant(strPropertyName), &pUserProperty);
							}
							if (pUserProperty)
							{
								CComVariant m_v;
								pUserProperty->get_Value(&m_v);
								USES_CONVERSION;
								strRet = OLE2T(m_v.bstrVal);
							}
						}
					}
				}
				//CoUninitialize();
			});

			tasks.run_and_wait(task);
			//CString strRet = _T("");
			//if (pFolder)
			//{
			//	CComPtr<OutLook::_StorageItem> pStoreItem;
			//	pFolder->GetStorage(CComBSTR(strSubject), olIdentifyBySubject, &pStoreItem);
			//	if (pStoreItem)
			//	{
			//		long nSize = 0;
			//		pStoreItem->get_Size(&nSize);
			//		CComPtr<OutLook::UserProperties> pUserProperties;
			//		pStoreItem->get_UserProperties(&pUserProperties);
			//		CComPtr<OutLook::UserProperty> pUserProperty;
			//		if (nSize == 0)
			//		{
			//			//return strRet;
			//		}
			//		else
			//		{
			//			pUserProperties->Item(CComVariant(strPropertyName), &pUserProperty);
			//		}
			//		if (pUserProperty)
			//		{
			//			CComVariant m_v;
			//			pUserProperty->get_Value(&m_v);
			//			strRet = OLE2T(m_v.bstrVal);
			//		}
			//	}
			//}
			return strRet;
		}

		void COutLookCloudAddin::AddFormPageToInspector(_Inspector* pInspector, CString strPageName, CString strInfoXml, BOOL bShowInSpector, BOOL bSetCurrentPage)
		{
			if (strPageName == _T(""))
				strPageName = _T("Tangram");
			if (pInspector)
			{
				OutLook::Pages* pPages = nullptr;
				pInspector->get_ModifiedFormPages((IDispatch**)&pPages);
				if (pPages)
				{
					CComPtr<IDispatch> pPage;
					CComPtr<IDispatch> pItem;
					CComVariant vName(strPageName);
					pPages->Item(vName, &pItem);
					if (pItem == nullptr)
					{
						HRESULT hr = pPages->Add(vName, &pPage);
						if (pPage)
						{
							::VariantClear(&vName);
							BSTR bstrPageName = strPageName.AllocSysString();
							if(bShowInSpector)
								hr = pInspector->ShowFormPage(bstrPageName);
							if(bSetCurrentPage)
								hr = pInspector->SetCurrentFormPage(bstrPageName);

							AddTangramPropertyToInspector(pInspector, strPageName, strInfoXml);
						}
					}
				}
			}
		}

		STDMETHODIMP COutLookCloudAddin::TangramCommand(IDispatch* RibbonControl)
		{
			if (m_spRibbonUI)
				m_spRibbonUI->Invalidate();
			CComBSTR bstrTag(L"");
			CComBSTR bstrID(L"");
			CComQIPtr<IRibbonControl> pIRibbonControl(RibbonControl);
			if (pIRibbonControl)
			{
				pIRibbonControl->get_Id(&bstrID);
				pIRibbonControl->get_Tag(&bstrTag);
			}

			CString strTag = OLE2T(bstrTag);
			int nPos = strTag.Find(_T("TangramButton.Cmd"));
			strTag.Replace(_T("TangramButton.Cmd."), _T(""));
			int nCmdIndex = _wtoi(strTag);
			HWND hWnd = ::GetActiveWindow();
			auto it = m_mapInspector.find(hWnd);
			if (it != m_mapInspector.end())
			{
				it->second->TangramCommand(nCmdIndex);
			}
			else
			{
				auto it = m_mapOutLookPlusExplorerMap.find(hWnd);
				if (it != m_mapOutLookPlusExplorerMap.end())
				{
					it->second->TangramCommand(nCmdIndex);
				}
			}
			return S_OK;
		}

		HRESULT COutLookCloudAddin::OnConnection(IDispatch* pHostApp, int ConnectMode)
		{
			if (m_pApplication)
				return S_OK;
			HWND hWnd = ::GetActiveWindow();
			HWND _hWnd = hWnd;
			wchar_t user[256] = { 0 };
			ULONG size = _countof(user);
			if (!GetUserNameEx(NameDisplay, user, &size))
			{
				// GetUserNameEx may fail (for example on Home editions). use the login name
				size = _countof(user);
				GetUserName(user, &size);
			}
			m_strUser = user;
			pHostApp->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pApplication);

			CTangramXmlParse m_Parse;
			if (m_Parse.LoadXml(theApp.m_strConfigFile))
			{
				CTangramXmlParse* pUIParse = m_Parse.GetChild(_T("OutLookUI"));
				if (pUIParse)
				{
					int nCount = pUIParse->GetCount();
					for (int i = 0; i < nCount; i++)
					{
						CTangramXmlParse* pChild = pUIParse->GetChild(i);
						m_mapUIXML[pChild->name()] = pChild->xml();
					}
				}
				CString strItems = m_Parse.attr(_T("OutLookItems"), _T(""));
				if (strItems != _T(""))
				{
					int nPos = strItems.Find(_T(","));
					int nIndex = 1;
					while (nPos != -1)
					{
						m_mapItemName[nIndex] = strItems.Left(nPos);
						nIndex++;
						strItems = strItems.Mid(nPos + 1);
						nPos = strItems.Find(_T(","));
					}
				}
			}

			CComPtr<OutLook::_NameSpace> pSession;
			m_pApplication->get_Session(&pSession);
			pSession->get_Stores(&m_pStores);
			long nCount;
			m_pStores->get_Count(&nCount);
			::PostMessage(m_hHostWnd, WM_TANGRAMMSG, 0, 0);
			//HRESULT hr = ((COutLookApplicationEvents*)this)->DispEventAdvise(m_pApplication);
			//hr = ((COutLookApplicationEvents_10*)this)->DispEventAdvise(m_pApplication);
			HRESULT hr = ((COutLookApplicationEvents_11*)this)->DispEventAdvise(m_pApplication);
			//hr = ((COutLookStoresEvents_12*)this)->DispEventAdvise(m_pStores);
			m_pApplication->get_Explorers(&m_pExplorers);
			if (m_pExplorers)
			{
				hr = ((COutLookExplorersEvents*)this)->DispEventAdvise(m_pExplorers);
				m_pApplication->get_Inspectors(&m_pInspectors);
				hr = ((COutLookInspectorsEvents*)this)->DispEventAdvise(m_pInspectors);
			}
			CComPtr<IDispatch> pDisp;
			m_pApplication->ActiveWindow(&pDisp);
			if (pDisp)
			{
				CComQIPtr<_Explorer> pExplorer(pDisp);
				if (pExplorer)
				{
					m_strCurrentKey = _T("Microsoft.Outlook.Explorer");
					COutLookExplorer* pOutlookExplorer = new COutLookExplorer();
					pOutlookExplorer->m_strKey = m_strCurrentKey;
					m_mapOutLookPlusExplorerMap[hWnd] = pOutlookExplorer;
					pOutlookExplorer->m_pExplorer = pExplorer;
					pOutlookExplorer->m_hWnd = _hWnd;
					m_pActiveOutlookExplorer = pOutlookExplorer;
					if (pOutlookExplorer->m_pExplorer)
					{
						hr = ((COutLookExplorerEvents*)pOutlookExplorer)->DispEventAdvise(pOutlookExplorer->m_pExplorer);
						hr = ((COutLookExplorerEvents_10*)pOutlookExplorer)->DispEventAdvise(pOutlookExplorer->m_pExplorer);
						if (m_nVer > 0x000f)
						{
							hr = pOutlookExplorer->m_pExplorer->get_AccountSelector(&pOutlookExplorer->m_pAccountSelector);
							if (hr == S_OK)
							{
								pOutlookExplorer->m_pAccountSelector->AddRef();
								hr = ((COutLookAccountSelectorEvents*)pOutlookExplorer)->DispEventAdvise(pOutlookExplorer->m_pAccountSelector);
							}
						}
						pOutlookExplorer->m_pExplorer->get_NavigationPane(&pOutlookExplorer->m_pNavigationPane);
						hr = ((COutLookNavigationPaneEvents_12*)pOutlookExplorer)->DispEventAdvise(pOutlookExplorer->m_pNavigationPane);
						hr = pOutlookExplorer->m_pNavigationPane->get_CurrentModule(&pOutlookExplorer->m_pCurrentModule);
						if (hr == S_OK)
						{
							pOutlookExplorer->m_pCurrentModule->AddRef();

							pOutlookExplorer->m_pCurrentModule->get_NavigationModuleType(&pOutlookExplorer->m_nNType);
							CComBSTR bstrName("");
							switch (pOutlookExplorer->m_nNType)
							{
							case olModuleMail:
							{
								CComQIPtr<_MailModule> _pModule(pOutlookExplorer->m_pCurrentModule);
								_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
							}
							break;
							case olModuleCalendar:
							{
								CComQIPtr<_CalendarModule> _pModule(pOutlookExplorer->m_pCurrentModule);
								_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
							}
							break;
							case olModuleContacts:
							{
								CComQIPtr<_ContactsModule> _pModule(pOutlookExplorer->m_pCurrentModule);
								_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
							}
							break;
							case olModuleTasks:
							{
								CComQIPtr<_TasksModule> _pModule(pOutlookExplorer->m_pCurrentModule);
								_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
							}
							break;
							case olModuleJournal:
							{
								CComQIPtr<_JournalModule> _pModule(pOutlookExplorer->m_pCurrentModule);
								_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
							}
							break;
							case olModuleNotes:
							{
								CComQIPtr<_NotesModule> _pModule(pOutlookExplorer->m_pCurrentModule);
								_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
							}
							break;
							case olModuleFolderList:
							case olModuleShortcuts:
							case olModuleSolutions:
								break;
							}
							if (pOutlookExplorer->m_pNavigationGroups)
							{
								pOutlookExplorer->m_pNavigationGroups->AddRef();
								hr = ((COutLookNavigationGroupsEvents_12*)pOutlookExplorer)->DispEventAdvise(pOutlookExplorer->m_pNavigationGroups);
							}
						}
					}
				}
				else
				{
					CComQIPtr<_Inspector> pInspector(pDisp);
					if (pInspector)
					{
						hWnd = ::FindWindowEx(hWnd, NULL, _T("AfxWndW"), NULL);
						hWnd = ::FindWindowEx(hWnd, NULL, _T("AfxWndW"), NULL);
						CWindow m_Wnd;
						m_Wnd.Attach(hWnd);
						hWnd = m_Wnd.GetWindow(GW_CHILD).m_hWnd;
						m_Wnd.Detach();
						hWnd = ::GetDlgItem(hWnd, 0x0103f);
						m_hWwbWnd = ::FindWindowEx(hWnd, NULL, _T("_WwB"), NULL);
						IDispatch* pDisp;
						pInspector->get_CurrentItem(&pDisp);
						COutLookInspector* pOutLookPlusItemWindow = nullptr;
						CComQIPtr<_MailItem> m_pMailItem(pDisp);
						pOutLookPlusItemWindow = new COutLookInspector();
						pOutLookPlusItemWindow->m_strKey = m_strCurrentKey;
						m_pCurOpenItem = pOutLookPlusItemWindow;
						hr = ((COutLookInspectorEvents*)pOutLookPlusItemWindow)->DispEventAdvise(pDisp);
						hr = ((COutLookInspectorEvents_10*)pOutLookPlusItemWindow)->DispEventAdvise(pDisp);
						hr = ((COutLookItemEvents*)pOutLookPlusItemWindow)->DispEventAdvise(pDisp);
						//hr = ((COutLookItemEvents_10*)pOutLookPlusItemWindow)->DispEventAdvise(pDisp);
						pOutLookPlusItemWindow->m_pDisp = pDisp;
						//pOutLookPlusItemWindow->AddRef();

						m_hPWwbWnd = NULL;
					}
				}
			}
			return S_OK;
		}

		HRESULT COutLookCloudAddin::OnDisconnection(int DisConnectMode)
		{
			((COutLookExplorersEvents*)this)->DispEventUnadvise(m_pExplorers);
			((COutLookInspectorsEvents*)this)->DispEventUnadvise(m_pInspectors);
			//HRESULT hr = ((COutLookApplicationEvents*)this)->DispEventUnadvise(m_pApplication);
			//hr = ((COutLookApplicationEvents_10*)this)->DispEventUnadvise(m_pApplication);
			HRESULT hr = ((COutLookApplicationEvents_11*)this)->DispEventUnadvise(m_pApplication);
			//hr = ((COutLookStoresEvents_12*)this)->DispEventUnadvise(m_pStores);
			if (m_pTangramInspectorItems)
			{
				m_pTangramInspectorItems->DispEventUnadvise(m_pTangramInspectorItems->m_pItems);
				m_pTangramInspectorItems->m_pItems->Release();
				delete m_pTangramInspectorItems;
			}
			m_pApplication.p->Release();
			m_pApplication.Detach();
			return S_OK;
		}

		CString COutLookCloudAddin::ExportTangramData()
		{
			CString strRet = _T("");
			CComPtr<OutLook::_NameSpace> pNameSpace;
			m_pApplication->get_Session(&pNameSpace);
			if (pNameSpace)
			{
				CComPtr<_Store> pDefaultStore;
				pNameSpace->get_DefaultStore(&pDefaultStore);
				CComPtr<_Stores> pStores;
				pNameSpace->get_Stores(&pStores);
				if (pStores)
				{
					long nCount = -1;
					pStores->get_Count(&nCount);
					if (nCount)
					{
						for (int i = 1; i <= nCount; i++)
						{
							CComVariant index(i);
							CComPtr<_Store> pStore;
							pStores->Item(index, &pStore);
							if (pStore)
							{
								CComPtr<MAPIFolder> pRootFolder;
								pStore->GetRootFolder(&pRootFolder);
								if (pRootFolder)
								{
									CComPtr<_Folders> pFolders;
									pRootFolder->get_Folders(&pFolders);
								}
							}
						}
					}
				}
			}
			return strRet;
		}

		CString COutLookCloudAddin::GetDefaultFolderXml(_Store* pStore, OlDefaultFolders m_folderenum)
		{
			CString strData = _T("");
			CComPtr<MAPIFolder> pDefaultFolder;
			pStore->GetDefaultFolder(m_folderenum, &pDefaultFolder);
			if (pDefaultFolder)
			{
				CComBSTR bstrEntryID(L"");
				pDefaultFolder->get_EntryID(&bstrEntryID);
				CString strKey = OLE2T(bstrEntryID);
				strData.Format(_T("<Tangram%s><FolderType>%d</FolderType></Tangram%s>"), strKey, m_folderenum, strKey);
			}
			return strData;
		}

		STDMETHODIMP COutLookCloudAddin::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml)
		{
			/*
			Microsoft.Outlook.Explorer
			Microsoft.Outlook.Mail.Read
			Microsoft.Outlook.Mail.Compose
			Microsoft.Outlook.MeetingRequest.Read
			Microsoft.Outlook.MeetingRequest.Send
			Microsoft.Outlook.Appointment
			Microsoft.Outlook.Contact
			Microsoft.Outlook.Journal
			Microsoft.Outlook.Task
			Microsoft.Outlook.DistributionList
			Microsoft.Outlook.Report
			Microsoft.Outlook.Resend
			Microsoft.Outlook.Post.Read
			Microsoft.Outlook.Post.Compose
			Microsoft.Outlook.Sharing.Read
			Microsoft.Outlook.Sharing.Compose

			enum OlItemType
			{
			olMailItem = 0,
			olAppointmentItem = 1,
			olContactItem = 2,
			olTaskItem = 3,
			olJournalItem = 4,
			olNoteItem = 5,
			olPostItem = 6,
			olDistributionListItem = 7,
			olMobileItemSMS = 11,
			olMobileItemMMS = 12
			};
			*/
			CString strRet = _T("");
			m_strCurrentKey = OLE2T(RibbonID);
			auto it = m_mapUIXML.find(m_strCurrentKey);
			if (it != m_mapUIXML.end())
			{
				CString strXml = it->second;
				int nPos = strXml.Find(_T("<customUI"));
				if (nPos != -1)
				{
					strRet = strXml.Mid(nPos);
					nPos = strRet.Find(_T("</customUI>"));
					if (nPos != -1)
					{
						strRet = strRet.Left(nPos + 11);
					}
				}
			}

			if (strRet != _T(""))
			{
				*RibbonXml = strRet.AllocSysString();
				return S_OK;
			}

			return E_POINTER;
		}

		void COutLookCloudAddin::WindowCreated(LPCTSTR lpszClass, LPCTSTR strName, HWND hPWnd, HWND hWnd)
		{
		}

		void COutLookCloudAddin::WindowDestroy(HWND hWnd)
		{
			::GetClassName(hWnd, theApp.m_szBuffer, MAX_PATH);
			CString strClassName = CString(theApp.m_szBuffer);
			if (strClassName == _T("ThunderDFrame"))
			{
				//OnDestroyVbaForm(hWnd);
				auto it = m_mapVBAForm.find(hWnd);
				if (it != m_mapVBAForm.end())
					m_mapVBAForm.erase(it);
			}
			else
			{
				OnCloseOfficeObj(strClassName, hWnd);
			}
		}

		//void COutLookCloudAddin::OnBeforeStoreRemove(_Store* Store, VARIANT_BOOL* Cancel)
		//{
		//}

		//void COutLookCloudAddin::OnStoreAdd(_Store* Store)
		//{
		//}

		//begin COutLookExplorersEvents:
		void COutLookCloudAddin::OnNewExplorer(_Explorer* Explorer)
		{
			COutLookExplorer* pOutlookExplorer = new COutLookExplorer();
			pOutlookExplorer->m_pExplorer.Attach(Explorer);
			Explorer->AddRef();
			pOutlookExplorer->m_strKey = _T("Microsoft.Outlook.Explorer");
			m_pActiveOutlookExplorer = pOutlookExplorer;

			HRESULT hr = ((COutLookExplorerEvents*)pOutlookExplorer)->DispEventAdvise(pOutlookExplorer->m_pExplorer);
			hr = ((COutLookExplorerEvents_10*)pOutlookExplorer)->DispEventAdvise(pOutlookExplorer->m_pExplorer);
			pOutlookExplorer->m_pExplorer->get_NavigationPane(&pOutlookExplorer->m_pNavigationPane);
			if (pOutlookExplorer->m_pNavigationPane)
			{
				hr = ((COutLookNavigationPaneEvents_12*)pOutlookExplorer)->DispEventAdvise(pOutlookExplorer->m_pNavigationPane);
				hr = pOutlookExplorer->m_pNavigationPane->get_CurrentModule(&pOutlookExplorer->m_pCurrentModule);
				if (hr == S_OK)
				{
					pOutlookExplorer->m_pCurrentModule->AddRef();
					pOutlookExplorer->m_pCurrentModule->get_NavigationModuleType(&pOutlookExplorer->m_nNType);
					CComBSTR bstrName("");
					switch (pOutlookExplorer->m_nNType)
					{
					case olModuleMail:
					{
						CComQIPtr<_MailModule> _pModule(pOutlookExplorer->m_pCurrentModule);
						_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
					}
					break;
					case olModuleCalendar:
					{
						CComQIPtr<_CalendarModule> _pModule(pOutlookExplorer->m_pCurrentModule);
						_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
					}
					break;
					case olModuleContacts:
					{
						CComQIPtr<_ContactsModule> _pModule(pOutlookExplorer->m_pCurrentModule);
						_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
					}
					break;
					case olModuleTasks:
					{
						CComQIPtr<_TasksModule> _pModule(pOutlookExplorer->m_pCurrentModule);
						_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
					}
					break;
					case olModuleJournal:
					{
						CComQIPtr<_JournalModule> _pModule(pOutlookExplorer->m_pCurrentModule);
						_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
					}
					break;
					case olModuleNotes:
					{
						CComQIPtr<_NotesModule> _pModule(pOutlookExplorer->m_pCurrentModule);
						_pModule->get_NavigationGroups(&pOutlookExplorer->m_pNavigationGroups);
					}
					break;
					case olModuleFolderList:
					case olModuleShortcuts:
						break;
					case olModuleSolutions:
					{
						CComQIPtr<_SolutionsModule> _pModule(pOutlookExplorer->m_pCurrentModule);
						VARIANT_BOOL bVisible;
						_pModule->get_Visible(&bVisible);
						if (bVisible == false)
						{
							_pModule->put_Visible(true);
						}
						else
						{
						}
					}
					break;
					}
					if (pOutlookExplorer->m_nNType == olModuleSolutions)
					{
						//if (pOutlookExplorer->m_hNavWnd&&theApp.m_pSolutionFrame)
						//{
						//	if (theApp.m_pSolutionHelperWnd&&::IsWindow(theApp.m_pSolutionHelperWnd->m_hWnd))
						//		::ShowWindow(theApp.m_pSolutionHelperWnd->m_hWnd, SW_SHOW);
						//	if (pOutlookExplorer->m_hButton)
						//		::ShowWindow(pOutlookExplorer->m_hButton, SW_HIDE);
						//	theApp.m_pSolutionFrame->Attach();
						//}
					}
					else
					{
						//if (pOutlookExplorer->m_hNavWnd&&theApp.m_pSolutionFrame)
						//{
						//	if (theApp.m_pSolutionHelperWnd&&::IsWindow(theApp.m_pSolutionHelperWnd->m_hWnd))
						//		::ShowWindow(theApp.m_pSolutionHelperWnd->m_hWnd, SW_HIDE);
						//	if (pOutlookExplorer->m_hButton)
						//		::ShowWindow(pOutlookExplorer->m_hButton, SW_SHOW);
						//	theApp.m_pSolutionFrame->Deatch();
						//}
					}
					if (pOutlookExplorer->m_pNavigationGroups)
					{
						pOutlookExplorer->m_pNavigationGroups->AddRef();
						hr = ((COutLookNavigationGroupsEvents_12*)pOutlookExplorer)->DispEventAdvise(pOutlookExplorer->m_pNavigationGroups);
					}
				}
			}

			if (m_nVer>0x000f)
			{
				hr = pOutlookExplorer->m_pExplorer->get_AccountSelector(&pOutlookExplorer->m_pAccountSelector);
				if (hr == S_OK)
				{
					pOutlookExplorer->m_pAccountSelector->AddRef();
					hr = ((COutLookAccountSelectorEvents*)pOutlookExplorer)->DispEventAdvise(pOutlookExplorer->m_pAccountSelector);
				}
			}
			::PostMessage(m_hHostWnd, WM_TANGRAMNEWOUTLOOKOBJ, 1, (LPARAM)pOutlookExplorer);
		}
		//end COutLookExplorersEvents

		//begin COutLookInspectorsEvents:
		void COutLookCloudAddin::OnNewInspector(_Inspector* Inspector)
		{
			IDispatch* pDisp;
			Inspector->get_CurrentItem(&pDisp);
			COutLookInspector* pOutLookPlusItemWindow = new COutLookInspector();
			BSTR szMember = ::SysAllocString(L"Class");
			HRESULT hr = S_OK;
			if (m_nGetClassDispID == -1)
			{
				hr = pDisp->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_USER_DEFAULT, &m_nGetClassDispID);
			}
			if (pDisp&&m_nGetClassDispID!=-1)
			{
				DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
				VARIANT result = { 0 };
				EXCEPINFO excepInfo;
				memset(&excepInfo, 0, sizeof excepInfo);
				UINT nArgErr = (UINT)-1;
				hr = pDisp->Invoke(m_nGetClassDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &result, &excepInfo, &nArgErr);
				if (S_OK == hr )
				{
					pOutLookPlusItemWindow->m_OlObjectClass = (OlObjectClass)result.intVal;
				}
				::VariantClear(&result);
			}
			::SysFreeString(szMember);

			m_pCurOpenItem = pOutLookPlusItemWindow;
			pOutLookPlusItemWindow->m_pInspector = Inspector;
			pOutLookPlusItemWindow->m_pInspector->AddRef();
			hr = ((COutLookItemEvents*)pOutLookPlusItemWindow)->DispEventAdvise(pDisp);
			hr = ((COutLookInspectorEvents*)pOutLookPlusItemWindow)->DispEventAdvise(Inspector);
			hr = ((COutLookInspectorEvents_10*)pOutLookPlusItemWindow)->DispEventAdvise(Inspector);
			//hr = ((COutLookItemEvents_10*)pOutLookPlusItemWindow)->DispEventAdvise(pDisp);
			pOutLookPlusItemWindow->m_pDisp = pDisp;
		}
		//end COutLookInspectorsEvents

		//New COutLookExplorer Object:
		COutLookExplorer::COutLookExplorer(void)
		{
			m_hWnd = NULL;
			m_strKey = _T("");
			m_hTaskPaneWnd				= NULL;
			m_hTaskPaneChildWnd			= NULL;
			m_hOutLookToday				= NULL;
			m_hTaskPane					= NULL;
			m_pTaskPane					= nullptr;
			m_pModuleDisp				= nullptr;
			m_pOnlineItem				= nullptr;
			m_pCurrentModule			= nullptr;
			m_pAccountSelector			= nullptr;
			m_pNavigationGroups			= nullptr;

			m_pPage						= nullptr;
			m_pFrame					= nullptr;
			m_pNewFolder				= nullptr;
			m_pTaskPanePage				= nullptr;
			m_pTaskPaneFrame			= nullptr;
			m_pInspectorContainerWnd	= nullptr;
		}

		COutLookExplorer::~COutLookExplorer(void)
		{
		}

		void COutLookExplorer::TangramCommand(int nIndex)
		{
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			switch (nIndex)
			{
			case 100:
			{
				pAddin->CreateCommonDesignerToolBar();
				CComPtr<MAPIFolder> pFolder;
				m_pExplorer->get_CurrentFolder(&pFolder);
				if (pFolder)
				{
					CComPtr<MAPIFolder> _pFolder;
					CComBSTR bstrEntryID(L"");
					CComBSTR bstrStoreID(L"");
					pFolder->get_EntryID(&bstrEntryID);
					pAddin->m_strDesignExplorerEntryID = OLE2T(bstrEntryID);
					pFolder->get_StoreID(&bstrStoreID);
					pAddin->m_strDesignExplorerStoreID = OLE2T(bstrStoreID);
					CComPtr<_NameSpace> pSession;
					HRESULT hr = pAddin->m_pApplication->get_Session(&pSession);
					pSession->GetFolderFromID(bstrEntryID, CComVariant(bstrStoreID), &_pFolder);
					CComPtr<_Explorer> pExplorer;
					hr = pAddin->m_pExplorers->Add(CComVariant(_pFolder), OlFolderDisplayMode::olFolderDisplayNoNavigation, &pExplorer);
					if (hr == S_OK)
					{
						pExplorer->Display();
						CString strKey = pAddin->m_strDesignExplorerEntryID;
						strKey += _T(",");
						strKey += pAddin->m_strDesignExplorerStoreID;
						pAddin->m_mapDesignExplorer[strKey] = pExplorer;
						::PostMessage(pAddin->m_hHostWnd, WM_TANGRAMMSG, 1963, 1222);
					}
				}
			}
			break;
			case 101:
			case 102:
			{
				CComVariant varView;
				m_pExplorer->get_CurrentView(&varView);
				if (varView.vt == VT_DISPATCH)
				{
					CComQIPtr<OutLook::View> pView(varView.pdispVal);
					if (pView)
					{
						OlViewType viewType;
						pView->get_ViewType(&viewType);
						switch (viewType)
						{
						case OlViewType::olBusinessCardView:
							break;
						case OlViewType::olCalendarView:
							break;
						case OlViewType::olCardView:
							break;
						case OlViewType::olDailyTaskListView:
							break;
						case OlViewType::olIconView:
							break;
						case OlViewType::olPeopleView:
							break;
						case OlViewType::olTableView:
							{
								CComPtr<OutLook::MAPIFolder> pFolder;
								m_pExplorer->get_CurrentFolder(&pFolder);
								if (pFolder)
								{
								}
							}
							break;
						case OlViewType::olTimelineView:
							{
								CComQIPtr<_TimelineView> _pTimelineView(pView);
								if (_pTimelineView)
								{
									//_pTimelineView->
								}
							}
							break;
						}
					}
				}
			}
			break;
			}
		}

		//begin COutLookExplorerEvents:
		void COutLookExplorer::OnActivate()
		{
			if (m_hWnd == NULL)
				m_hWnd = ::GetActiveWindow();
			//if (::IsWindow(m_hWnd))
			//{
			//	::SetParent(theApp.m_pSolutionHelperWnd->m_hWnd, m_hWnd);
			//	theApp.m_pSolutionFrame->ModifyHost((long)m_pTangramSolutionHostWnd->m_hWnd, (LONGLONG)m_hNavWnd);
			//	if (m_nNType != olModuleSolutions)
			//	{
			//		theApp.m_pSolutionFrame->Deatch();
			//		::ShowWindow(m_hButton, SW_SHOW);
			//		::ShowWindow(theApp.m_pSolutionHelperWnd->m_hWnd, SW_HIDE);
			//	}
			//	else
			//	{
			//		::ShowWindow(m_hButton, SW_HIDE);
			//		::ShowWindow(theApp.m_pSolutionHelperWnd->m_hWnd, SW_SHOW);
			//	}
			//}
			//theApp.m_pActiveOutlookExplorer = this;
			//BSTR bstrHtml = L"";
			//CComPtr<IDispatch> pDoc;
			//	m_pExplorer->get_HTMLDocument(&pDoc);
			//	CComQIPtr<IHTMLDocument2> pDoc2(pDoc);
			//	if (pDoc2)
			//	{
			//		CComPtr<IHTMLElement>pBody;
			//		pDoc2->get_body(&pBody);
			//		pBody->get_outerHTML(&bstrHtml);
			//	}
		}

		void COutLookExplorer::OnFolderSwitch()
		{
			if (m_hWnd == NULL)
				m_hWnd = ::GetActiveWindow();
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			CComPtr<OutLook::MAPIFolder> _pFolder;
			m_pExplorer->get_CurrentFolder(&_pFolder);
			CString strXml = _T("");
			CString strPath = _T("");
			//IStream* pStream = 0;
			//HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_MAPIFolder, _pFolder.p, &pStream);
			//auto task = create_task([pStream,&strXml,&strPath]
			//{
			//	MAPIFolder* pFolder = NULL;
			//	HRESULT hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_MAPIFolder, (LPVOID *)&pFolder);
			//	BSTR bstrID = L"";
			//	pFolder->get_FolderPath(&bstrID);
			//	CComPtr<OutLook::_StorageItem> pStoreItem;
			//	hr = pFolder->GetStorage(CComBSTR(_T("Tangram")), olIdentifyBySubject, &pStoreItem);
			//	if (pStoreItem)
			//	{
			//		long nSize = 0;
			//		pStoreItem->get_Size(&nSize);
			//		CComPtr<OutLook::UserProperties> pUserProperties;
			//		pStoreItem->get_UserProperties(&pUserProperties);
			//		CComPtr<OutLook::UserProperty> pUserProperty;
			//		if (nSize)
			//		{
			//			pUserProperties->Item(CComVariant(_T("TangramProperties")), &pUserProperty);
			//		}
			//		if (pUserProperty)
			//		{
			//			CComVariant m_v;
			//			pUserProperty->get_Value(&m_v);
			//			//strXml = OLE2T(m_v.bstrVal);
			//		}
			//	}
			//}).wait();

			HWND _hwnd = GetCurrentFolderFrameHandle();
			if (_pFolder)
			{
				CString strXml=pAddin->GetFolderPropertyFromStore(_pFolder, _T("Tangram"), _T("TangramProperties"));
				if (strXml != _T(""))
				{
					IWndFrame* pFrame = nullptr;
					theApp.m_pHostCore->GetWndFrame((long)_hwnd, &pFrame);
					if (pFrame == nullptr)
					{
						IWndPage* pPage = nullptr;
						theApp.m_pHostCore->CreateWndPage((long)m_hOutLookToday, &pPage);
						if (::PathFileExists(theApp.m_strAppPath+_T("outlooktoday.appxml")))
							strXml = _T("outlooktoday.appxml");
						if (pPage)
							pPage->CreateFrame(CComVariant(0), CComVariant((long)_hwnd), CComBSTR(L"OutLook"), &pFrame);
					}
					if (pFrame)
					{
						CComBSTR bstrName(L"");
						_pFolder->get_FolderPath(&bstrName);
						IWndNode* pNode = nullptr;
						pFrame->Extend(bstrName, strXml.AllocSysString(), &pNode);
					}
				}
			}
			
			if (m_pNewFolder)
			{
				CComBSTR bstrScript(L"");
				HRESULT hr;
				if (pAddin->m_pTangramInspectorItems)
				{
					hr = pAddin->m_pTangramInspectorItems->DispEventUnadvise(pAddin->m_pTangramInspectorItems->m_pItems);
					hr = pAddin->m_pTangramInspectorItems->m_pItems->Release();
				}

				CComQIPtr<OutLook::MAPIFolder> pFolder(m_pNewFolder);
				if (pFolder)
				{
					pFolder->get_Name(&bstrScript);
					CComPtr<_Items> pItems;
					pFolder->get_Items(&pItems);
					if (pAddin->m_pTangramInspectorItems == nullptr)
						pAddin->m_pTangramInspectorItems = new CInspectorItems;
					pAddin->m_pTangramInspectorItems->m_pItems = pItems.p;
					pAddin->m_pTangramInspectorItems->m_pItems->AddRef();
					hr = pAddin->m_pTangramInspectorItems->DispEventAdvise(pItems.p);
					long nCount = 0;
					pItems->get_Count(&nCount);
					if (nCount == 0 && m_pInspectorContainerWnd)
					{
						if (m_pInspectorContainerWnd->m_pFrame)
						{
							HWND hwnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), L"", WS_CHILD, 0, 0, 0, 0, theApp.m_pHostCore->m_hHostWnd, NULL, AfxGetInstanceHandle(), NULL);
							HWND hChildWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), L"", WS_CHILD, 0, 0, 0, 0, (HWND)hwnd, NULL, AfxGetInstanceHandle(), NULL);
							m_pInspectorContainerWnd->m_pFrame->ModifyHost((long)hChildWnd);
							::DestroyWindow(hwnd);
							m_pInspectorContainerWnd->m_pFrame = nullptr;
							m_pInspectorContainerWnd->m_pPage = nullptr;
						}
					}
				}
			}
		}

		void COutLookExplorer::OnBeforeFolderSwitch(IDispatch* NewFolder, VARIANT_BOOL* Cancel)
		{
			CComBSTR bstrScript(L"");
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			
			CComQIPtr<OutLook::MAPIFolder> pFolder(NewFolder);
			if (pFolder)
			{
				if (m_pNewFolder)
				{
					m_pNewFolder->Release();
				}
				m_pNewFolder = NewFolder;
				m_pNewFolder->AddRef();
			}
		}

		void COutLookExplorer::OnViewSwitch()
		{
			CComVariant varView;
			m_pExplorer->get_CurrentView(&varView);
			if (varView.vt == VT_DISPATCH)
			{
				CComQIPtr<OutLook::View> pView(varView.pdispVal);
				if (pView)
				{
					OlViewType viewType;
					pView->get_ViewType(&viewType);
					switch (viewType)
					{
					case OlViewType::olBusinessCardView:
						break;
					case OlViewType::olCalendarView:
						break;
					case OlViewType::olCardView:
						break;
					case OlViewType::olDailyTaskListView:
						break;
					case OlViewType::olIconView:
						break;
					case OlViewType::olPeopleView:
						break;
					case OlViewType::olTableView:
						break;
					case OlViewType::olTimelineView:
						{
							CComQIPtr<_TimelineView> _pTimelineView(pView);
							if (_pTimelineView)
							{
								//_pTimelineView->
							}
						}
						break;
					}
				}
			}
		}

		void COutLookExplorer::OnBeforeViewSwitch(VARIANT NewView, VARIANT_BOOL* Cancel)
		{
			CComBSTR bstrViewName = NewView.bstrVal;
		}

		void COutLookExplorer::OnDeactivate()
		{
		}

		void COutLookExplorer::OnSelectionChange()
		{
			//CComPtr<OutLook::Selection> pSelection;
			//m_pExplorer->get_Selection(&pSelection);
			//if (pSelection)
			//{
			//	long nCount = 0;
			//	pSelection->get_Count(&nCount);
			//	if (nCount)
			//	{
			//		CComPtr<IDispatch> pDisp;
			//		pSelection->Item(CComVariant(1), &pDisp);
			//		if (pDisp)
			//		{
			//			_Inspector* pIns = pAddin->GetInspector(pDisp);
			//			if (pIns)
			//			{
			//				CComBSTR bstrCap(L"");
			//				pIns->get_Caption(&bstrCap);
			//			}
			//		}
			//	}
			//}
		}

		void COutLookExplorer::OnClose()
		{
			if (m_pNewFolder)
			{
				m_pNewFolder->Release();
				m_pNewFolder = nullptr;
			}
			HRESULT hr = ((COutLookNavigationPaneEvents_12*)this)->DispEventUnadvise(m_pNavigationPane);
			hr = ((COutLookExplorerEvents*)this)->DispEventUnadvise(m_pExplorer);
			hr = ((COutLookExplorerEvents_10*)this)->DispEventUnadvise(m_pExplorer);
			//if (theApp.m_nVer>0x000f)
			hr = ((COutLookAccountSelectorEvents*)this)->DispEventUnadvise(m_pAccountSelector);

			if (m_pNavigationGroups)
				hr = ((COutLookNavigationGroupsEvents_12*)this)->DispEventUnadvise(m_pNavigationGroups);
			if (m_pCurrentModule)
				m_pCurrentModule->Release();
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			auto it = pAddin->m_mapOutLookPlusExplorerMap.find(m_hWnd);
			if (it != pAddin->m_mapOutLookPlusExplorerMap.end())
			{
				pAddin->m_mapOutLookPlusExplorerMap.erase(it);
			}
		}
		//end COutLookExplorerEvents

		//begin COutLookAccountSelectorEvents:
		void COutLookExplorer::OnSelectedAccountChange(Account* SelectedAccount)
		{
		}
		//end COutLookAccountSelectorEvents

		//begin COutLookExplorerEvents_10:
		//void COutLookExplorer::OnBeforeMaximize(VARIANT_BOOL* Cancel)
		//{
		//}

		//void COutLookExplorer::OnBeforeMinimize(VARIANT_BOOL* Cancel)
		//{
		//}

		//void COutLookExplorer::OnBeforeMove(VARIANT_BOOL* Cancel)
		//{
		//}

		//void COutLookExplorer::OnBeforeSize(VARIANT_BOOL* Cancel)
		//{
		//}

		//void COutLookExplorer::OnBeforeItemCopy(VARIANT_BOOL* Cancel)
		//{
		//}

		//void COutLookExplorer::OnBeforeItemCut(VARIANT_BOOL* Cancel)
		//{
		//}

		//void COutLookExplorer::OnBeforeItemPaste(VARIANT_BOOL* Cancel)
		//{
		//}

		void COutLookExplorer::OnAttachmentSelectionChange()
		{
		}

		void COutLookExplorer::OnInlineResponse(IDispatch* Item)
		{
			//if (theApp.m_hPWwbWnd)
			//{
			//	CString m_strUIScript = _T("");
			//	map<int, CString>::iterator it = theApp.m_mapItemName.find(OutLookPlusMailItem);
			//	if (it != theApp.m_mapItemName.end())
			//	{
			//		CString strKey = it->second;
			//		if (strKey != _T(""))
			//		{
			//			map<CString, CString>::iterator it2 = theApp.m_mapUIScript.find(strKey);
			//			if (it2 == theApp.m_mapUIScript.end())
			//			{
			//				CTangramXmlParse m_Parse;
			//				if (m_Parse.LoadXml(theApp.m_strConfigFile))
			//				{
			//					CTangramXmlParse* pParse = m_Parse.GetChild(strKey);
			//					if (pParse)
			//					{
			//						CTangramXmlParse* pParse2 = pParse->GetChild(_T("UIScript"));
			//						if (pParse2)
			//						{
			//							m_strUIScript = pParse2->xml();
			//							theApp.m_mapUIScript[strKey] = m_strUIScript;
			//						}
			//					}
			//				}
			//			}
			//			else
			//			{
			//				m_strUIScript = it2->second;
			//			}
			//		}
			//	}
			//m_pOnlineItem = Item;
			//	m_pOnlineItem->AddRef();

			//	CComPtr<ITangramWindow> pWnd;
			//	CString strUrl = m_strUIScript;
			//	strUrl.Replace(_T("%s"), theApp.m_strUser);
			//	theApp.m_pTangramManager->AddToolBarForWnd((LONGLONG)theApp.m_hPWwbWnd, (LONGLONG)theApp.m_hWwbWnd, (IDispatch*)m_pOnlineItem, CComBSTR(strUrl), &m_pOnLineItemHostWindow);
			//	m_pOnLineItemHostWindow->AddRef();
			//	m_pOnLineItemHostWindow->put_Extender(m_pOnlineItem);
			//	theApp.m_hPWwbWnd = NULL;
			//}
		}

		void COutLookExplorer::OnInlineResponseClose()
		{
			//if (m_pOnLineItemHostWindow)
			//{
			//	if (m_pOnlineItem)
			//	{
			//		m_pOnlineItem->Release();
			//		m_pOnlineItem = NULL;
			//	}
			//	m_pOnLineItemHostWindow->CloseFrame();
			//}
		}
		//end dCOutLookExplorerEvents_10

		//begin COutLookNavigationPaneEvents_12:
		void COutLookExplorer::OnModuleSwitch(NavigationModule* pCurrentModule)
		{
			CComQIPtr<_NavigationModule> _pNavigationModule((_NavigationModule*)pCurrentModule);

			if (_pNavigationModule)
			{
				if (m_pNavigationGroups)
				{
					HRESULT hr = ((COutLookNavigationGroupsEvents_12*)this)->DispEventUnadvise(m_pNavigationGroups);
					m_pNavigationGroups->Release();
					m_pNavigationGroups = nullptr;
					m_pCurrentModule->Release();
					m_pCurrentModule = _pNavigationModule.p;
					m_pCurrentModule->AddRef();
				}
				HRESULT hr = _pNavigationModule->get_NavigationModuleType(&m_nNType);
				CComBSTR bstrName("");
				switch (m_nNType)
				{
				case olModuleMail:
				{
					CComQIPtr<_MailModule> _pModule(m_pCurrentModule);
					if (_pModule)
						_pModule->get_NavigationGroups(&m_pNavigationGroups);
				}
				break;
				case olModuleCalendar:
				{
					CComQIPtr<_CalendarModule> _pModule(m_pCurrentModule);
					if (_pModule)
						_pModule->get_NavigationGroups(&m_pNavigationGroups);
				}
				break;
				case olModuleContacts:
				{
					CComQIPtr<_ContactsModule> _pModule(m_pCurrentModule);
					if (_pModule)
						_pModule->get_NavigationGroups(&m_pNavigationGroups);
				}
				break;
				case olModuleTasks:
				{
					CComQIPtr<_TasksModule> _pModule(m_pCurrentModule);
					if (_pModule)
						_pModule->get_NavigationGroups(&m_pNavigationGroups);
				}
				break;
				case olModuleJournal:
				{
					CComQIPtr<_JournalModule> _pModule(m_pCurrentModule);
					if (_pModule)
						_pModule->get_NavigationGroups(&m_pNavigationGroups);
				}
				break;
				case olModuleNotes:
				{
					CComQIPtr<_NotesModule> _pModule(m_pCurrentModule);
					if (_pModule)
						_pModule->get_NavigationGroups(&m_pNavigationGroups);
				}
				break;
				case olModuleFolderList:
				case olModuleShortcuts:
				case olModuleSolutions:
					break;
				}
				if (m_nNType == olModuleSolutions)
				{
					//if (::IsChild(m_hWnd, theApp.m_pSolutionHelperWnd->m_hWnd) == false)
					//{
					//	::SetParent(theApp.m_pSolutionHelperWnd->m_hWnd, m_hWnd);
					//	theApp.m_pSolutionFrame->ModifyHost((LONGLONG)m_pTangramSolutionHostWnd->m_hWnd, (LONGLONG)m_hNavWnd);
					//}
					//if (m_hNavWnd&&theApp.m_pSolutionFrame)
					//{
					//	if (theApp.m_pSolutionHelperWnd&&::IsWindow(theApp.m_pSolutionHelperWnd->m_hWnd))
					//		::ShowWindow(theApp.m_pSolutionHelperWnd->m_hWnd, SW_SHOW);
					//	if (m_hButton)
					//		::ShowWindow(m_hButton, SW_HIDE);
					//	theApp.m_pSolutionFrame->Attach();
					//}
				}
				else
				{
					//if (m_hNavWnd&&theApp.m_pSolutionFrame)
					//{
					//	if (theApp.m_pSolutionHelperWnd&&::IsWindow(theApp.m_pSolutionHelperWnd->m_hWnd))
					//		::ShowWindow(theApp.m_pSolutionHelperWnd->m_hWnd, SW_HIDE);
					//	if (m_hButton)
					//		::ShowWindow(m_hButton, SW_SHOW);
					//	theApp.m_pSolutionFrame->Deatch();
					//}
				}
				if (m_pNavigationGroups)
				{
					m_pNavigationGroups->AddRef();
					hr = ((COutLookNavigationGroupsEvents_12*)this)->DispEventAdvise(m_pNavigationGroups);
				}
			}
		}
		//end COutLookNavigationPaneEvents_12

		//begin COutLookNavigationGroupsEvents_12:
		void COutLookExplorer::OnSelectedChange(NavigationFolder* NavigationFolder)
		{
			_NavigationFolder* _pNavigationFolder = nullptr;
		}

		void COutLookExplorer::OnNavigationFolderAdd(NavigationFolder* NavigationFolder)
		{
		}

		void COutLookExplorer::OnNavigationFolderRemove()
		{
		}
		//end COutLookNavigationGroupsEvents_12

		HWND COutLookExplorer::GetCurrentFolderFrameHandle()
		{
			CString strID = _T("Tangram");
			CComBSTR bstrEntryID(L"");
			CComPtr<MAPIFolder> pFolder;
			m_pExplorer->get_CurrentFolder(&pFolder);
			if (pFolder)
			{
				CComBSTR bstrName(L"");
				pFolder->get_Name(&bstrName);
				pFolder->get_EntryID(&bstrEntryID);
				strID += OLE2T(bstrEntryID);
				auto it = m_mapFolderWnd.find(strID);
				if (it != m_mapFolderWnd.end())
					return it->second;
			}

			HWND hRet = NULL;
			HWND hWnd = ::GetActiveWindow();
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			BOOL bHaveView = false;
				
			CComVariant m_v;
			m_pExplorer->get_CurrentView(&m_v);
			if (m_v.vt == VT_DISPATCH)
			{
				bHaveView = true;
				CComQIPtr<OutLook::View>pView(m_v.pdispVal);
				if (pView)
				{
					OlViewType viewType;
					pView->get_ViewType(&viewType);
					switch (viewType)
					{
					case OlViewType::olBusinessCardView:
						break;
					case OlViewType::olCalendarView:
						break;
					case OlViewType::olCardView:
						break;
					case OlViewType::olDailyTaskListView:
						break;
					case OlViewType::olIconView:
						break;
					case OlViewType::olPeopleView:
					case OlViewType::olTableView:
					{
						if (m_pInspectorContainerWnd == nullptr)
						{
							HWND _hWnd = ::FindWindowEx(hWnd, NULL, _T("rctrl_renwnd32"), NULL);
							HWND hWnd2 = _hWnd;
							if (_hWnd)
							{
								_hWnd = ::FindWindowEx(_hWnd, NULL, _T("AfxWndW"), NULL);
								if (_hWnd)
								{
									m_pInspectorContainerWnd = new CInspectorContainerWnd();
									m_pInspectorContainerWnd->SubclassWindow(_hWnd);
								}
							}
						}
						if (m_pInspectorContainerWnd)
						{
							return m_pInspectorContainerWnd->m_hWnd;
						}
					}
					break;
					case OlViewType::olTimelineView:
					{
						CComQIPtr<_TimelineView> _pTimelineView(pView);
						if (_pTimelineView)
						{
							//_pTimelineView->
						}
					}
					break;
					}
				}
			}
			else
			{
				m_hOutLookToday = ::FindWindowEx(m_hWnd, NULL, _T("Shell Embedding"), NULL);
				if (m_hOutLookToday)
				{
					hRet = ::FindWindowEx(m_hOutLookToday, NULL, _T("Shell DocObject View"), NULL);
				}
			}

			return hRet;
		}

		void COutLookExplorer::SetDesignState()
		{
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			HWND hWnd = ::GetActiveWindow();
			HWND _hwnd = NULL;
			CInspectorContainerWnd* pWnd = nullptr;
			CComVariant varView;
			m_pExplorer->get_CurrentView(&varView);
			CComPtr<MAPIFolder> pFolder;
			m_pExplorer->get_CurrentFolder(&pFolder);
			CString strXml = pAddin->GetFolderPropertyFromStore(pFolder, _T("Tangram"), _T("TangramProperties"));

			if (varView.vt == VT_DISPATCH)
			{
				CComQIPtr<OutLook::View> pView(varView.pdispVal);
				if (pView)
				{
					OlViewType viewType;
					pView->get_ViewType(&viewType);
					switch (viewType)
					{
					case OlViewType::olBusinessCardView:
						break;
					case OlViewType::olCalendarView:
						break;
					case OlViewType::olCardView:
						break;
					case OlViewType::olDailyTaskListView:
						break;
					case OlViewType::olIconView:
						break;
					case OlViewType::olPeopleView:
					case OlViewType::olTableView:
					{
						_hwnd = ::FindWindowEx(hWnd, NULL, _T("rctrl_renwnd32"), NULL);
						if (m_pInspectorContainerWnd == NULL)
						{
							if (_hwnd)
							{
								_hwnd = ::FindWindowEx(_hwnd, NULL, _T("AfxWndW"), NULL);
								if (_hwnd)
								{
									m_pInspectorContainerWnd = new CInspectorContainerWnd();
									m_pInspectorContainerWnd->SubclassWindow(_hwnd);
								}
							}
						}
						pWnd = m_pInspectorContainerWnd;
					}
					break;
					case OlViewType::olTimelineView:
					{
						CComQIPtr<_TimelineView> _pTimelineView(pView);
						if (_pTimelineView)
						{
						}
					}
					break;
					}
				}
			}
			else
			{
				m_hOutLookToday = ::FindWindowEx(m_hWnd, NULL, _T("Shell Embedding"), NULL);
				if (m_hOutLookToday)
				{
					_hwnd = ::FindWindowEx(m_hOutLookToday, NULL, _T("Shell DocObject View"), NULL);
					pWnd = new CInspectorContainerWnd();
					pWnd->SubclassWindow(_hwnd);
				}
			}
			pWnd->m_pOutLookExplorer = this;
			IWndFrame* pFrame = nullptr;
			theApp.m_pHostCore->GetWndFrame((long)pWnd->m_hWnd, &pWnd->m_pFrame);
			if (pWnd->m_pFrame == nullptr)
			{
				theApp.m_pHostCore->CreateWndPage((long)::GetParent(pWnd->m_hWnd), &pWnd->m_pPage);
				if (pWnd->m_pPage)
					pWnd->m_pPage->CreateFrame(CComVariant(0), CComVariant((long)pWnd->m_hWnd), CComBSTR(L""), &pWnd->m_pFrame);
			}
			if (pWnd->m_pFrame)
			{
				CComBSTR bstrName(L"");
				pFolder->get_FolderPath(&bstrName);
				IWndNode* pNode = nullptr;
				if(strXml==_T(""))
					strXml  = _T("<Tangram><window><node name=\"Start\" /></window></Tangram>");
				pWnd->m_pFrame->Extend(bstrName, strXml.AllocSysString(), &pNode);
				CWndFrame* _pFrame = (CWndFrame*)pWnd->m_pFrame;
				theApp.HostPosChanged(_pFrame->m_pBindingNode);
				_pFrame->m_bDesignerState = true;
				pAddin->CreateCommonDesignerToolBar();
			}
		}

		CInspectorItem::CInspectorItem(void)
		{
			m_strXml					= _T("");
			m_pItem						= nullptr;
			m_pInspectorContainerWnd	= nullptr;
		}

		CInspectorItem::~CInspectorItem(void)
		{
		}

		void CInspectorItem::OnRead()
		{
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			m_strXml = pAddin->GetTangramPropertyFromItem(m_pItem, _T("Tangram"));
		}

		void CInspectorItem::OnUnload()
		{
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			long nKey = (long)this;
			auto it = pAddin->m_mapTangramInspectorItem.find(nKey);
			if (it != pAddin->m_mapTangramInspectorItem.end())
				pAddin->m_mapTangramInspectorItem.erase(it);
			HRESULT hr = ((COutLookItemEvents*)this)->DispEventUnadvise(m_pItem);
			//hr = ((COutLookItemEvents_10*)this)->DispEventUnadvise(m_pItem);
			ULONG dw = m_pItem->Release();
			while(dw)
				dw = m_pItem->Release();
			delete this;
		}

		CInspectorItems::CInspectorItems(void)
		{
			m_pItems = nullptr;
		}

		CInspectorItems::~CInspectorItems(void)
		{

		}

		void CInspectorItems::OnItemAdd(IDispatch* Item) 
		{
		}

		void CInspectorItems::OnItemChange(IDispatch* Item) 
		{
		}

		void CInspectorItems::OnItemRemove() 
		{
		}

		//New COutLookInspector Object:
		COutLookInspector::COutLookInspector(void)
		{
			m_bReadOnly				= false;
			m_strKey				= _T("");
			m_strUIScript			= _T("");
			m_hWnd					= nullptr;
			m_hHostWnd				= NULL;
			m_hTaskPaneWnd			= NULL;
			m_hTaskPaneChildWnd		= NULL;
			m_pTaskPane				= nullptr;
			m_pInspector			= nullptr;
			m_pTaskPaneFrame		= nullptr;
			m_pTaskPanePage			= nullptr;
			m_pActiveOutLookPage	= nullptr;
		}

		COutLookInspector::~COutLookInspector(void)
		{
		}

		void COutLookInspector::OnPageChange(BSTR ActivePageName) 
		{
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			m_strActivePageName = OLE2T(ActivePageName);
			m_strActivePageName.Trim();
			m_strActivePageName.Replace(_T(" "), _T("_"));
			::PostMessage(pAddin->m_hHostWnd, WM_TANGRAMACTIVEINSPECTORPAGE, (WPARAM)this, 0);
		}

		void COutLookInspector::OnActivate()
		{
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			HWND hActive = ::GetActiveWindow();
			auto it = pAddin->m_mapOutLookPlusExplorerMap.find(hActive);
			if (it == pAddin->m_mapOutLookPlusExplorerMap.end())
			{
				auto it2 = pAddin->m_mapInspector.find(hActive);
				if (it2 == pAddin->m_mapInspector.end())
				{
					m_hHostWnd = hActive;
					pAddin->m_mapInspector[m_hHostWnd] = this;
				}

				if (m_pActiveOutLookPage)
				{
					m_pActiveOutLookPage->ActivePage();
				}
			}
		}

		void COutLookInspector::OnDeactivate() 
		{
		}

		void COutLookInspector::OnClose() 
		{
			if (m_hHostWnd)
			{
				HRESULT hr = S_FALSE;
				COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
				if (m_pTaskPanePage)
				{
					HWND _hWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), L"", WS_CHILD, 0, 0, 0, 0, theApp.m_pHostCore->m_hHostWnd, NULL, AfxGetInstanceHandle(), NULL);
					HWND _hChildWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), L"", WS_CHILD, 0, 0, 0, 0, (HWND)_hWnd, NULL, AfxGetInstanceHandle(), NULL);
					if (::IsWindow(m_hWnd))
					{
						m_pTaskPaneFrame->ModifyHost((long)_hChildWnd);
						::DestroyWindow(m_hWnd);
					}
				}
				hr = ((COutLookInspectorEvents*)this)->DispEventUnadvise(m_pInspector);
				hr = ((COutLookInspectorEvents_10*)this)->DispEventUnadvise(m_pInspector);
				hr = ((COutLookItemEvents*)this)->DispEventUnadvise(m_pDisp);
				//hr = ((COutLookItemEvents_10*)this)->DispEventUnadvise(m_pDisp);
				ULONG dw = m_pDisp->Release();
				while(dw)
					dw = m_pDisp->Release();
				m_pDisp = nullptr;
				auto it = pAddin->m_mapInspector.find(m_hHostWnd);
				if (it != pAddin->m_mapInspector.end())
				{
					pAddin->m_mapInspector.erase(it);
				}
			}
			delete this;
		}

		void COutLookInspector::TangramCommand(int nIndex)
		{
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			switch (nIndex)
			{
			case 100:
			{
				CString strXml = pAddin->GetTangramPropertyFromInspector(m_pInspector, _T("Tangram"));
				if (strXml == _T(""))
				{
					CString str = theApp.m_strExeName;
					CString strCaption = _T("");
					str += _T("\\");
					switch (m_OlObjectClass)
					{
					case olMail:
						{
							str += _T("mail");
							strCaption = _T("Please select Mail Template:");
						}
						break;
					case olAppointment:
						{
							str += _T("Appointment");
							strCaption = _T("Please select Appointment Template:");
						}
						break;
					case olTask:
						{
							str += _T("Task");
							strCaption = _T("Please select Task Template:");
						}
						break;
					case olContact:
						{
							str += _T("Contact");
							strCaption = _T("Please select Contact Template:");
						}
						break;
					case olJournal:
						{
							str += _T("Journal");
							strCaption = _T("Please select Journal Template:");
						}
						break;
					case olPost:
						{
							str += _T("Post");
							strXml = pAddin->GetDocTemplateXml(_T("Please select Post Template:"), str);
						}
						break;
					case olDistributionList:
						{
							str += _T("DistributionList");
							strCaption = _T("Please select DistributionList Template:");
						}
						break;
					}
					strXml = pAddin->GetDocTemplateXml(strCaption, str);
				}
				if (m_pActiveOutLookPage)
					m_pActiveOutLookPage->DesignPage();
			}
			break;
			case 101:
			case 102:
			{
				auto iter = pAddin->m_mapTaskPaneMap.find((LONG)m_hHostWnd);
				if (iter != pAddin->m_mapTaskPaneMap.end())
					pAddin->m_mapTaskPaneMap[(LONG)m_hHostWnd]->put_Visible(true);
				else
				{
					CString strKey = _T("TaskPaneUI");
					CString strXml = m_strUIScript;
					if(strXml==_T(""))
						strXml = _T("<TaskPaneUI><window><node name=\"Start\" /></window></TaskPaneUI>");
					CTangramXmlParse m_Parse;
					if (m_Parse.LoadXml(strXml))
					{
						CTangramXmlParse* pParse = m_Parse.GetChild(strKey);
						if (pParse)
						{
							strXml = pParse->xml();
						}
						else
							strXml = _T("<TaskPaneUI><window><node name=\"Start\" /></window></TaskPaneUI>");
						if (strXml != _T(""))
						{
							{
								VARIANT vWindow;
								vWindow.vt = VT_DISPATCH;
								vWindow.pdispVal = nullptr;
								Office::_CustomTaskPane* _pCustomTaskPane;
								CString strCap = _T("");
								if (pParse)
								{
									CTangramXmlParse* pNodeParse = m_Parse.FindItem(_T("node"));
									if(pNodeParse)
										strCap = pNodeParse->attr(_T("caption"), _T(""));
								}
								if (strCap == _T(""))
									strCap = _T("Start");
								CComBSTR bstrCap(strCap);
								HRESULT hr = pAddin->m_pCTPFactory->CreateCTP(L"Tangram.Ctrl.1", bstrCap, vWindow, &_pCustomTaskPane);
								_pCustomTaskPane->AddRef();
								_pCustomTaskPane->put_Visible(true);
								pAddin->m_mapTaskPaneMap[(LONG)m_hHostWnd] = _pCustomTaskPane;
								CComPtr<ITangramCtrl> pCtrlDisp;
								_pCustomTaskPane->get_ContentControl((IDispatch**)&pCtrlDisp);
								if (pCtrlDisp)
								{
									long hWnd = 0;
									pCtrlDisp->get_HWND(&hWnd);
									HWND hPWnd = ::GetParent((HWND)hWnd);
									m_hTaskPane = (HWND)hWnd;
									CWindow m_Wnd;
									m_Wnd.Attach(hPWnd);
									m_Wnd.ModifyStyle(0, WS_CLIPSIBLINGS | WS_CLIPCHILDREN);
									m_Wnd.Detach();

									if (m_pTaskPanePage == nullptr)
									{
										HRESULT hr = theApp.m_pHostCore->CreateWndPage((long)hPWnd, &m_pTaskPanePage);
										if (m_pTaskPanePage)
										{
											m_pTaskPanePage->CreateFrame(CComVariant(0), CComVariant((long)hWnd), CComBSTR(L"TaskPane"), &m_pTaskPaneFrame);
											if (m_pTaskPaneFrame)
											{
												IWndNode* pNode = nullptr;
												m_pTaskPaneFrame->Extend(CComBSTR("Default"), strXml.AllocSysString(), &pNode);
											}
										}
									}
									else
										m_pTaskPaneFrame->ModifyHost(hWnd);
								}
							}
						}
					}
				}
				if (nIndex == 102)
				{
					
				}
			}
			break;
			case 103:
			{
				if (m_pInspector)
				{
					OutLook::Pages* pPages = nullptr;
					m_pInspector->get_ModifiedFormPages((IDispatch**)&pPages);
					long nCount = 0;
					pPages->get_Count(&nCount);
					if (nCount<5)
					{
						COutLookCustomizeFormDlg m_Dlg;
						m_Dlg.m_pOutLookInspector = this;
						m_Dlg.DoModal();
					}
					else
					{

					}
				}
			}
				break;
			case 104:
				break;
			}
		}

		//begin COutLookItemEvents:
		void COutLookInspector::ActivePage()
		{
			HWND hActive = ::GetActiveWindow();
			if (hActive == NULL)
				return;
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			BOOL bInExplorer = false;
			auto it = pAddin->m_mapOutLookPlusExplorerMap.find(hActive);
			if (it != pAddin->m_mapOutLookPlusExplorerMap.end())
			{
				bInExplorer = true;
				if (m_hHostWnd == NULL)
				{
					m_hHostWnd = ::FindWindowEx(hActive, NULL, _T("rctrl_renwnd32"), NULL);
					if (m_hHostWnd == NULL)
						return;
				}
			}
			else
			{
				m_hHostWnd = hActive;
			}
			if (m_hWnd == NULL)
				m_hWnd = ::FindWindowEx(m_hHostWnd, NULL, _T("AfxWndW"), NULL);
			if (m_hWnd == NULL)
				return;
			HWND hwnd = NULL;
			if (bInExplorer == false)
				hwnd = ::FindWindowEx(m_hWnd, NULL, _T("AfxWndW"), NULL);
			else
			{
				hwnd = m_hWnd;
			}
			auto it2= m_mapOutLookPage.find(m_strActivePageName);
			if (it2 != m_mapOutLookPage.end())
			{
				m_pActiveOutLookPage = it2->second;
				if (m_pActiveOutLookPage->m_bDesignState)
				{
					pAddin->CreateCommonDesignerToolBar();
					CWndFrame* pFrame = (CWndFrame*)m_pActiveOutLookPage->m_pFrame;
					if (theApp.m_pDesigningFrame != pFrame)
					{
						theApp.m_pDesigningFrame = pFrame;
						pFrame->UpdateDesignerTreeInfo();
					}
				}
			}
			else
			{
				CString strXml = pAddin->GetTangramPropertyFromItem(m_pDisp, _T("Tangram"));
				//if (strXml == _T(""))
				//{
				//	switch (m_OlObjectClass)
				//	{
				//		//olAppointment = 26,
				//		//	olMeetingRequest = 53,
				//		//	olMeetingCancellation = 54,
				//		//	olMeetingResponseNegative = 55,
				//		//	olMeetingResponsePositive = 56,
				//		//	olMeetingResponseTentative = 57,
				//			//olContact = 40,
				//		//	olDocument = 41,
				//		//	olJournal = 42,
				//		//	olMail = 43,
				//		//	olNote = 44,
				//		//	olPost = 45,
				//		//	olReport = 46,
				//		//	olRemote = 47,
				//		//	olTask = 48,
				//		//	olTaskRequest = 49,
				//		//	olTaskRequestUpdate = 50,
				//		//	olTaskRequestAccept = 51,
				//		//	olTaskRequestDecline = 52,
				//	case OlObjectClass::olMail:
				//	{
				//		//CString str = theApp.m_strExeName;
				//		//str += _T("\\Mail");
				//		//CString strTemplate = pAddin->GetDocTemplateXml(str);
				//		//CTangramXmlParse m_Parse;
				//		//bool bLoad = m_Parse.LoadXml(strTemplate);
				//		//if (bLoad == false)
				//		//	bLoad = m_Parse.LoadFile(strTemplate);
				//		//if (bLoad == false)
				//		//	return;
				//	}
				//		break;
				//	case OlObjectClass::olTask:
				//		break;
				//	}
				//}
				m_pActiveOutLookPage = new COutLookPageWnd();
				m_pActiveOutLookPage->m_pOutLookInspector = this;
				m_pActiveOutLookPage->m_strName = m_strActivePageName;
				m_pActiveOutLookPage->m_strXml = theApp.GetXmlData(m_strActivePageName, strXml);
				OutLook::Pages* pPages = nullptr;
				m_pInspector->get_ModifiedFormPages((IDispatch**)&pPages);
				CComPtr<IDispatch> pItem;
				pPages->Item(CComVariant(m_strActivePageName), &pItem);
				if (pItem)
				{
					CComQIPtr<MSForm::_UserForm> pForm(pItem);
					CComQIPtr<IOleWindow> pOleWnd(pForm);
					if(pOleWnd)
					{
						pOleWnd->GetWindow(&hwnd);
						m_pActiveOutLookPage->m_hFrameWnd = ::GetParent(hwnd);
						m_pActiveOutLookPage->SubclassWindow(::GetParent(m_pActiveOutLookPage->m_hFrameWnd));
						m_pActiveOutLookPage->m_pForm = pForm.p;
						m_pActiveOutLookPage->m_pForm->AddRef();
					}
				}
				else
				{
					m_pActiveOutLookPage->SubclassWindow(hwnd);
					m_pActiveOutLookPage->m_hFrameWnd = GetWindow(hwnd, GW_CHILD);
				}
				m_mapOutLookPage[m_strActivePageName] = m_pActiveOutLookPage;
			}
			if (bInExplorer == false)
			{
				m_pActiveOutLookPage->ActivePage();
			}
		}

		void COutLookInspector::OnOpen(VARIANT_BOOL* Cancel)
		{
		};

		void COutLookInspector::OnCustomAction(IDispatch* Action, IDispatch* Response, VARIANT_BOOL* Cancel)
		{
			CComQIPtr<OutLook::Action> pItem(Action);
		};

		void COutLookInspector::OnCustomPropertyChange(BSTR Name)
		{
		};

		void COutLookInspector::OnForward(IDispatch* Forward, VARIANT_BOOL* Cancel)
		{
			//×ª·¢
			CComQIPtr<_MailItem> pItem(Forward);
			if (pItem)
			{
				//CComBSTR bstrSubject(L"");
				//pItem->get_Subject(&bstrSubject);
				//pItem->get_HTMLBody(&bstrSubject);
				//* Cancel = true;
			}
		};

		void COutLookInspector::OnClose(VARIANT_BOOL* Cancel)
		{
		};

		void COutLookInspector::OnPropertyChange(BSTR Name)
		{
		};

		void COutLookInspector::OnRead()
		{
		};

		void COutLookInspector::OnReply(IDispatch* Response, VARIANT_BOOL* Cancel)
		{
			CComQIPtr<_MailItem> pItem(Response);
			if (pItem)
			{
				//CComBSTR bstrSubject(L"");
				//pItem->get_Subject(&bstrSubject);
				//pItem->get_HTMLBody(&bstrSubject);
				//* Cancel = true;
			}
		};

		void COutLookInspector::OnReplyAll(IDispatch* Response, VARIANT_BOOL* Cancel)
		{
			CComQIPtr<_MailItem> pItem(Response);
			if (pItem)
			{
				//CComBSTR bstrSubject(L"");
				//pItem->get_Subject(&bstrSubject);
				//pItem->get_HTMLBody(&bstrSubject);
				//* Cancel = false;
			}
		};

		void COutLookInspector::OnSend(VARIANT_BOOL* Cancel)
		{
		};

		void COutLookInspector::OnWrite(VARIANT_BOOL* Cancel)
		{
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			IStream* pStream = 0;
			HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID__Inspector, m_pInspector, &pStream);

			auto task = create_task([pStream, pAddin, this]()
			{
				CoInitializeEx(NULL, COINIT_MULTITHREADED);
				_Inspector* pInspector = nullptr;
				HRESULT hr = ::CoGetInterfaceAndReleaseStream(pStream, IID__Inspector, (LPVOID *)&pInspector);
				if (pInspector)
				{
					CString strPageData = _T("");
					CString strKeys = _T(",");
					CString strXml = _T("");
					theApp.Lock();
					for (auto it : m_mapOutLookPage)
					{
						CWndFrame* pFrame = (CWndFrame*)it.second->m_pFrame;
						if (pFrame)
						{
							strKeys += it.first;
							strKeys += _T(",");
							CWndNode* pNode = pFrame->m_pWorkNode;
							pAddin->UpdateWndNode(pNode);
							strPageData += pNode->m_pTangramParse->xml();
						}
					}
					strXml = pAddin->GetTangramPropertyFromInspector(pInspector, _T("Tangram"));
					theApp.Unlock();
					CString strData = _T("");
					CString strPagesXml = theApp.GetXmlData(_T("OutLookPages"), strXml);
					strData = _T("<OutLookPages>");
					if (strPagesXml != _T(""))
					{
						CString strKey = _T("<window>");
						CString _strKey = _T("");
						CString strName = _T("");
						int nPos = strPagesXml.Find(strKey);
						while (nPos != -1)
						{
							strName = strPagesXml.Left(nPos);
							nPos = strName.ReverseFind('<');
							strName = strName.Mid(nPos + 1);
							nPos = strName.ReverseFind('>');
							strName = strName.Left(nPos);
							strName.Trim();
							_strKey = _T(",");
							_strKey += strName;
							_strKey += _T(",");
							if (strKeys.Find(_strKey) == -1)
								strData += theApp.GetXmlData(strName, strPagesXml);
							strKey = _T("</window>");
							nPos = strPagesXml.Find(strKey);
							strPagesXml = strPagesXml.Mid(nPos + 9);
							nPos = strPagesXml.Find(_T("<window>"));
						}
					}
					strData += strPageData;
					strData += _T("</OutLookPages>");
					if (strData != _T("<OutLookPages></OutLookPages>"))
					{
						pAddin->AddTangramPropertyToInspector(pInspector, _T("Tangram"), strData);
					}
				}
				CoUninitialize();
			});
		};

		void COutLookInspector::OnBeforeCheckNames(VARIANT_BOOL* Cancel)
		{
		};

		void COutLookInspector::OnAttachmentAdd(Attachment* Attachment)
		{
		};

		void COutLookInspector::OnAttachmentRead(Attachment* Attachment)
		{
		};

		void COutLookInspector::OnBeforeAttachmentSave(Attachment* Attachment, VARIANT_BOOL* Cancel)
		{
		};
		//end COutLookItemEvents

		//begin COutLookItemEvents_10:
		//void COutLookInspector::OnBeforeDelete(IDispatch* Item, VARIANT_BOOL* Cancel)
		//{
		//};

		//void COutLookInspector::OnAttachmentRemove(Attachment* Attachment)
		//{
		//};

		//void COutLookInspector::OnBeforeAttachmentAdd(Attachment* Attachment, VARIANT_BOOL* Cancel)
		//{
		//};

		//void COutLookInspector::OnBeforeAttachmentPreview(Attachment* Attachment, VARIANT_BOOL* Cancel)
		//{
		//};

		//void COutLookInspector::OnBeforeAttachmentRead(Attachment* Attachment, VARIANT_BOOL* Cancel)
		//{
		//};

		//void COutLookInspector::OnBeforeAttachmentWriteToTempFile(Attachment* Attachment, VARIANT_BOOL* Cancel)
		//{
		//};

		//void COutLookInspector::OnUnload()
		//{
		//};

		//void COutLookInspector::OnBeforeAutoSave(VARIANT_BOOL* Cancel)
		//{
		//};

		//void COutLookInspector::OnBeforeRead()
		//{
		//};

		//void COutLookInspector::OnAfterWrite()
		//{
		//};

		//void COutLookInspector::OnReadComplete(VARIANT_BOOL* Cancel)
		//{
		//};
		//end COutLookItemEvents_10

		COutLookPageWnd::COutLookPageWnd(void) 
		{
			m_strXml			= _T("");
			m_bDesignState		= false;
			m_pPage				= nullptr;
			m_pForm				= nullptr;
			m_pFrame			= nullptr;
			m_pOutLookInspector = nullptr;
		}

		COutLookPageWnd::~COutLookPageWnd(void) 
		{
		}

		void COutLookPageWnd::OnFinalMessage(HWND hWnd)
		{
			CWindowImpl::OnFinalMessage(hWnd);
			delete this;
		}

		void COutLookPageWnd::ActivePage()
		{
			COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
			HWND hActive = ::GetActiveWindow();
			auto it = pAddin->m_mapOutLookPlusExplorerMap.find(hActive);
			if (it == pAddin->m_mapOutLookPlusExplorerMap.end())
			{
				m_pOutLookInspector->m_hHostWnd = hActive;
				pAddin->m_mapInspector[hActive] = m_pOutLookInspector;
				m_pOutLookInspector->m_hWnd = ::FindWindowEx(hActive, NULL, _T("AfxWndW"), NULL);
				HWND hWnd = NULL;
				if (m_pOutLookInspector->m_hWnd)
				{
					hWnd = ::FindWindowEx(m_pOutLookInspector->m_hWnd, NULL, _T("AfxWndW"), NULL);
					BOOL bVisible = ::IsWindowVisible(hWnd);
					while (hWnd&&bVisible == false)
					{
						HWND _hWnd = ::FindWindowEx(m_pOutLookInspector->m_hWnd, hWnd, _T("AfxWndW"), NULL);
						if (_hWnd)
						{
							hWnd = _hWnd;
							bVisible = ::IsWindowVisible(hWnd);
						}
						else
							break;
					}
					if (hWnd != m_hWnd)
					{
						UnsubclassWindow(true);
						SubclassWindow(hWnd);
					}
					HWND hFrame = ::GetWindow(hWnd, GW_CHILD);
					if (m_hFrameWnd != hFrame)
					{
						m_hFrameWnd = hFrame;
						if (m_pFrame)
						{
							m_pFrame->ModifyHost((long)hFrame);
						}
					}
				}
				if (m_strXml!=_T(""))
				{
					BSTR bstrName = m_strName.AllocSysString();
					if (m_pPage == nullptr)
					{
						theApp.m_pHostCore->CreateWndPage((long)m_hWnd, &m_pPage);
						if (m_pPage)
						{
							m_pPage->CreateFrame(CComVariant(0), CComVariant((long)m_hFrameWnd), bstrName, &m_pFrame);
						}
					}
					if (m_pFrame)
					{
						IWndNode* pNode = nullptr;
						m_pFrame->Extend(bstrName, m_strXml.AllocSysString(), &pNode);
						CWndNode* _pNode = ((CWndFrame*)m_pFrame)->m_pWorkNode;
						theApp.HostPosChanged(_pNode);
					}
					::SysFreeString(bstrName);
				}
			}
		}

		void COutLookPageWnd::DesignPage()
		{
			if (m_bDesignState == false)
			{
				m_bDesignState = true;
				if (m_pPage == nullptr)
				{
					theApp.m_pHostCore->CreateWndPage((long)m_hWnd, &m_pPage);
					if (m_pPage)
					{
						BSTR bstrName = m_strName.AllocSysString();
						m_pPage->CreateFrame(CComVariant(0), CComVariant((long)m_hFrameWnd), bstrName, &m_pFrame);
						if (m_pFrame)
						{
							m_pFrame->put_DesignerState(true);
							theApp.m_pHostCore->CreateCommonDesignerToolBar();
							IWndNode* pNode = nullptr;
							if (m_strXml == _T(""))
							{
								CString strName = m_strName;
								strName.Replace(_T(" "), _T("_"));
								m_strXml.Format(_T("<%s><window><node name=\"Start\" /></window></%s>"), strName, strName);
							}
							m_pFrame->Extend(bstrName, m_strXml.AllocSysString(), &pNode);
							theApp.m_pDesigningFrame = (CWndFrame*)m_pFrame;
							theApp.m_pDesigningFrame->m_bDesignerState = true;
							theApp.m_pDesigningFrame->UpdateDesignerTreeInfo();
						}
						::SysFreeString(bstrName);
					}
				}
			}
		}

		LRESULT COutLookPageWnd::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
		{
			if (m_pForm)
			{
				m_pForm->Release();
				m_pForm = nullptr;
			}

			return DefWindowProc(uMsg, wParam, lParam);
		}

		CInspectorContainerWnd::CInspectorContainerWnd()
		{
			m_strXml			= _T("");
			m_hFrameWnd			= NULL;
			m_pPage				= nullptr;
			m_pFrame			= nullptr;
			m_pOutLookExplorer	= nullptr;
		}

		CInspectorContainerWnd::~CInspectorContainerWnd()
		{
		}

		LRESULT CInspectorContainerWnd::OnTangramSave(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
		{
			if (m_pOutLookExplorer)
			{
				CComPtr<MAPIFolder> pFolder;
				MAPIFolder* _pFolder = nullptr;
				m_pOutLookExplorer->m_pExplorer->get_CurrentFolder(&pFolder);
				if (pFolder)
				{
					CWndFrame* pFrame = (CWndFrame*)m_pFrame;
					CComBSTR bstrXml(L"");
					COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
					pAddin->UpdateWndNode(pFrame->m_pWorkNode);
					pFrame->GetXml(_T(""), &bstrXml);
					CString strXml = theApp.GetXmlData(_T("Tangram"), OLE2T(bstrXml));
					pAddin->WriteFolderPropertyToStore(pFolder, _T("Tangram"), _T("TangramProperties"), strXml);
				}
			}
			return  DefWindowProc(uMsg, wParam, lParam);
		}

		LRESULT CInspectorContainerWnd::OnTangramMsg(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
		{
			if (m_pFrame&&wParam == 4096)
			{
				m_hFrameWnd = ::GetWindow(m_hWnd, GW_CHILD);
				m_pFrame->ModifyHost((long)m_hFrameWnd);
			}
			return  DefWindowProc(uMsg, wParam, lParam);
		}

		LRESULT CInspectorContainerWnd::OnItemLoad(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
		{
			if (m_pFrame)
			{
				HWND hwnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), _T(""), WS_CHILD, 0, 0, 0, 0, theApp.m_pHostCore->m_hHostWnd, NULL, AfxGetInstanceHandle(), NULL);
				m_pFrame->ModifyHost((long)::CreateWindowEx(NULL, _T("Tangram Window Class"), _T(""), WS_CHILD, 0, 0, 0, 0, (HWND)hwnd, NULL, AfxGetInstanceHandle(), NULL));
				::DestroyWindow(hwnd);
				m_pFrame = nullptr;
				m_pPage = nullptr;
			}
			if (m_strXml != _T(""))
			{
				CTangramXmlParse m_Parse;
				if (m_Parse.LoadXml(m_strXml))
				{
					m_hFrameWnd = ::GetWindow(m_hWnd, GW_CHILD);
					theApp.m_pTangram->CreateWndPage((long)m_hWnd, &m_pPage);
					m_pPage->CreateFrame(CComVariant(0), CComVariant((long)m_hFrameWnd), CComBSTR(L"Default"), &m_pFrame);
					if (m_pFrame)
					{
						IWndNode* pNode = nullptr;
						CTangramXmlParse* pChild = m_Parse.GetChild(0);
						m_pFrame->Extend(CComBSTR(L""), CComBSTR(pChild->xml()), &pNode);
					}
				}
			}

			return 0;// DefWindowProc(uMsg, wParam, lParam);
		}

		void CInspectorContainerWnd::OnFinalMessage(HWND hWnd)
		{
			CWindowImpl::OnFinalMessage(hWnd);
			delete this;
		}
	}
}
