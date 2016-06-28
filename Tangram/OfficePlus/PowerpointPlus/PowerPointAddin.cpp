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
* mailto:sunhuizlz@yeah.net
* http://www.CloudAddin.com
*
********************************************************************************/

#include "../../stdafx.h"
#include "../../CloudAddinApp.h"
#include "../../WndPage.h"
#include "../../WndNode.h"
#include "../../WndFrame.h"
#include "../../TangramHtmlTreeWnd.h"
#include "fm20.h"
#include "PowerPointAddin.h"
#include "PowerpointPlusEvents.cpp"

using namespace OfficeCloudPlus::PowerPointPlus;

namespace OfficeCloudPlus
{
	namespace PowerPointPlus
	{
		CPowerPntObject::CPowerPntObject(void)
		{
			m_hTaskPane = NULL;
			m_hTaskPaneWnd = NULL;
			m_hTaskPaneChildWnd = NULL;
			m_pPresentation = nullptr;
		}

		CPowerPntObject::~CPowerPntObject(void)
		{
		}

		void CPowerPntObject::OnObjDestory()
		{
			CPowerPntCloudAddin* pAddin = ((CPowerPntCloudAddin*)theApp.m_pHostCore);

			if (pAddin->m_pActivePowerPntObject == this)
			{
				pAddin->m_pActivePowerPntObject = nullptr;
				theApp.m_pWndNode = nullptr;
			}
			if (m_pPresentation)
			{
				CString strKey = m_pPresentation->m_strKey;
				auto it2 = m_pPresentation->find(m_hChildClient);
				if (it2 != m_pPresentation->end())
				{
					m_pPresentation->erase(it2);
				}
				size_t nCount = m_pPresentation->size();
				if (nCount > 0 && m_pPresentation->m_pFrame)
				{
					auto it2 = m_pPresentation->begin();
					m_pPresentation->m_pFrame->ModifyHost((long)it2->second->m_hClient);
					if (m_hTaskPane)
					{
						m_hTaskPane = NULL;
						m_pPresentation->m_pTaskPaneFrame->ModifyHost((long)it2->second->m_hTaskPaneChildWnd);
						::DestroyWindow(m_hTaskPaneWnd);
					}
				}
				if (nCount == 0)
				{
					CCloudAddinPresentation* pTangramPresentation = NULL;
					auto it = ((CPowerPntCloudAddin*)theApp.m_pHostCore)->m_mapTangramPres.find(strKey);
					if (it != ((CPowerPntCloudAddin*)theApp.m_pHostCore)->m_mapTangramPres.end())
						((CPowerPntCloudAddin*)theApp.m_pHostCore)->m_mapTangramPres.erase(it);
					delete m_pPresentation;
				}
			}

			auto it = pAddin->m_mapTaskPaneMap.find((long)m_hForm);
			if (it != pAddin->m_mapTaskPaneMap.end())
			{
				it->second->Delete();
				pAddin->m_mapTaskPaneMap.erase(it);
			}
			auto it3 = pAddin->m_mapOfficeObjects2.find(m_hForm);
			if (it3 != pAddin->m_mapOfficeObjects2.end())
				pAddin->m_mapOfficeObjects2.erase(it3);
		}

		CCloudAddinPresentation::CCloudAddinPresentation()
		{
			m_strKey = _T("");
			m_strTaskPaneXml = _T("");
			m_pPage = nullptr;
			m_pFrame = nullptr;
			m_pTaskPanePage = nullptr;
			m_pTaskPaneFrame = nullptr;
		}

		CCloudAddinPresentation::~CCloudAddinPresentation()
		{

		}

		CPowerPntCloudAddin::CPowerPntCloudAddin()
		{
			m_pActivePowerPntObject = nullptr;
			m_pCurrentSavingPresentation = nullptr;
		}

		STDMETHODIMP CPowerPntCloudAddin::AddDocXml(IDispatch* pDocdisp, BSTR bstrXml, BSTR bstrKey, VARIANT_BOOL* bSuccess)
		{
			CComQIPtr<PowerPoint::_Presentation> pDoc(pDocdisp);
			if (pDoc)
			{
				CComPtr<Office::_CustomXMLParts> pCustomXMLParts;
				pDoc->get_CustomXMLParts(&pCustomXMLParts);
				_AddDocXml(pCustomXMLParts.p, bstrXml, bstrKey, bSuccess);
			}
			return S_OK;
		}

		STDMETHODIMP CPowerPntCloudAddin::GetDocXmlByKey(IDispatch* pDocdisp, BSTR bstrKey, BSTR* pbstrXml)
		{
			CComQIPtr<PowerPoint::_Presentation> pDoc(pDocdisp);
			if (pDoc)
			{
				CComPtr<Office::_CustomXMLParts> pCustomXMLParts;
				pDoc->get_CustomXMLParts(&pCustomXMLParts);
				_GetDocXmlByKey(pCustomXMLParts.p, bstrKey, pbstrXml);
			}
			return S_OK;
		}

		void CPowerPntCloudAddin::OnPresentationBeforeSave(/*[in]*/ _Presentation * Pres,	/*[in,out]*/ VARIANT_BOOL * Cancel) 
		{
			CComBSTR bstrName(L"");
			Pres->get_FullName(&bstrName);
			CString strKey = OLE2T(bstrName);
			m_pCurrentSavingPresentation = nullptr;
			auto it = m_mapTangramPres.find(strKey);
			if (it != m_mapTangramPres.end())
				m_pCurrentSavingPresentation = it->second;

			if (m_pCurrentSavingPresentation)
			{
				CWndFrame* pFrame = m_pCurrentSavingPresentation->m_pFrame;
				if (pFrame)
				{
					pFrame->UpdateWndNode();
				}
				pFrame = m_pCurrentSavingPresentation->m_pTaskPaneFrame;
				if (pFrame)
				{
					pFrame->UpdateWndNode();
				}
			}
		}

		void CPowerPntCloudAddin::OnPresentationSave(/*[in]*/ _Presentation * Pres)
		{
			CComBSTR bstrName(L"");
			Pres->get_FullName(&bstrName);
			CString strNewKey = OLE2T(bstrName);
			if (m_pCurrentSavingPresentation&&m_pCurrentSavingPresentation->m_strKey != strNewKey)
			{
				auto it = m_mapTangramPres.find(strNewKey);
				if (it != m_mapTangramPres.end())
				{
					m_mapTangramPres.erase(it);
					m_mapTangramPres[strNewKey] = m_pCurrentSavingPresentation;
					m_pCurrentSavingPresentation->m_strKey = strNewKey;
					m_pCurrentSavingPresentation = nullptr;
				}
			}
		}
		
		void CPowerPntCloudAddin::UpdateOfficeObj(IDispatch* pObj, CString strXml, CString _strName)
		{
			CComQIPtr<MSForm::_UserForm> pForm(pObj);
			if (pForm)
			{
				CComPtr<_Presentation> pDoc;
				m_pApplication->get_ActivePresentation(&pDoc);
				if (pDoc)
				{
					CComPtr<Office::_CustomXMLParts> pCustomXMLParts;
					pDoc->get_CustomXMLParts(&pCustomXMLParts);
					CComBSTR bstrXml(L"");
					_GetDocXmlByKey(pCustomXMLParts.p, CComBSTR(L"Tangrams"), &bstrXml);
					CString _strXml = OLE2T(bstrXml);
					CString strName = _T("");
					CTangramXmlParse m_Parse;
					if (m_Parse.LoadXml(_strXml))
					{
						DISPID dispID = 0x80010000;
						DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
						VARIANT result = { 0 };
						EXCEPINFO excepInfo;
						memset(&excepInfo, 0, sizeof excepInfo);
						UINT nArgErr = (UINT)-1;
						HRESULT hr = pObj->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &result, &excepInfo, &nArgErr);
						if (S_OK == hr && VT_BSTR == result.vt)
						{
							CComBSTR bstrName = result.bstrVal;
							strName = OLE2T(bstrName);
							::VariantClear(&result);
						}

						CTangramXmlParse* pParse = m_Parse.GetChild(_T("userforms"));
						if (pParse)
						{
							CTangramXmlParse* pParse2 = pParse->GetChild(strName);
							CString strXml2 = _T("");
							if (pParse2)
							{
								CTangramXmlParse* pParse3 = pParse2->GetChild(_T("window"));
								CString strGUID = theApp.GetNewGUID();
								pParse2->RemoveNode(pParse3);
								pParse2->put_text(strGUID);
								strXml2 = m_Parse.xml();
								strXml2.Replace(strGUID, strXml);
							}
							else
							{
								strXml2.Format(_T("<%s>%s</%s>"), strName, strXml, strName);
								CTangramXmlParse m_Parse2;
								m_Parse2.LoadXml(strXml2);
								pParse->AddNode(&m_Parse2, _T(""));
								strXml2 = m_Parse.xml();
							}
							VARIANT_BOOL bRet = false;
							AddDocXml(pDoc, CComBSTR(strXml2), CComBSTR(L"Tangrams"), &bRet);
						}
					}
				}
				return;
			}
			CComQIPtr<_Presentation> pDoc(pObj);
			if (pDoc)
			{
				CComPtr<Office::_CustomXMLParts> pCustomXMLParts;
				pDoc->get_CustomXMLParts(&pCustomXMLParts);
				CComBSTR bstrXml(L"");
				_GetDocXmlByKey(pCustomXMLParts.p, CComBSTR(L"Tangrams"), &bstrXml);
				CString _strXml = OLE2T(bstrXml);
				CTangramXmlParse m_Parse;
				if (m_Parse.LoadXml(_strXml))
				{
					CTangramXmlParse* pParse = m_Parse.GetChild(_T("default"));
					if (pParse)
					{
						pParse = pParse->GetChild(_T("default"));
						if (pParse)
						{
							CTangramXmlParse* pParse2 = pParse->GetChild(_strName);
							if (pParse2)
							{
								CTangramXmlParse* pParse3 = pParse2->GetChild(_T("window"));
								CString strGUID = theApp.GetNewGUID();
								pParse2->RemoveNode(pParse3);
								pParse2->put_text(strGUID);
								CString str = m_Parse.xml();
								str.Replace(strGUID, strXml);
								//_strXml = str
								VARIANT_BOOL bRet = false;
								AddDocXml(pDoc, CComBSTR(str), CComBSTR(L"Tangrams"), &bRet);
							}
						}
					}
				}
			}
		}

		STDMETHODIMP CPowerPntCloudAddin::TangramCommand(IDispatch* RibbonControl)
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
			CComPtr<DocumentWindow> pWindow;
			m_pApplication->get_ActiveWindow(&pWindow);
			LONG hMainWnd = 0;
			pWindow->get_HWND(&hMainWnd);
			HWND hWnd = (HWND)hMainWnd;// ::GetActiveWindow();
			CPowerPntObject* pWnd = nullptr;
			LRESULT lRes = ::SendMessage(hWnd, WM_TANGRAMDATA, 0, 0);
			if (lRes)
			{
				pWnd = (CPowerPntObject*)lRes;
			}
			CCloudAddinPresentation* pTangramPresentation = pWnd->m_pPresentation;
			switch (nCmdIndex)
			{
			case 100:
			break;
			case 101:
			{
				auto iter = m_mapTaskPaneMap.find(hMainWnd);
				if (iter != m_mapTaskPaneMap.end())
				{
					m_mapTaskPaneMap[hMainWnd]->put_Visible(true);
				}
				else
				{
					CString strXml = pTangramPresentation->m_strTaskPaneXml;
					if (strXml != _T(""))
					{
						CTangramXmlParse m_Parse;
						if (m_Parse.LoadXml(strXml))
						{
							VARIANT vWindow;
							vWindow.vt = VT_DISPATCH;
							vWindow.pdispVal = nullptr;
							Office::_CustomTaskPane* _pCustomTaskPane;
							CString strCap = _T("");
							strCap = m_Parse.attr(_T("Title"), _T(""));
							if (strCap == _T(""))
								strCap = _T("Tangram for PowerPoint");
							CComBSTR bstrCap(strCap);
							HRESULT hr = m_pCTPFactory->CreateCTP(L"Tangram.Ctrl.1", bstrCap, vWindow, &_pCustomTaskPane);
							_pCustomTaskPane->AddRef();
							_pCustomTaskPane->put_Visible(true);
							m_mapTaskPaneMap[hMainWnd] = _pCustomTaskPane;
							CComPtr<ITangramCtrl> pCtrlDisp;
							_pCustomTaskPane->get_ContentControl((IDispatch**)&pCtrlDisp);
							if (pCtrlDisp)
							{
								long hWnd = 0;
								pCtrlDisp->get_HWND(&hWnd);
								HWND hPWnd = ::GetParent((HWND)hWnd);
								pWnd->m_hTaskPane = (HWND)hWnd;
								if (pTangramPresentation->m_pTaskPanePage == nullptr)
								{
									HRESULT hr = theApp.m_pHostCore->CreateWndPage((long)hPWnd, &pTangramPresentation->m_pTaskPanePage);
									if (pTangramPresentation->m_pTaskPanePage)
									{
										IWndFrame* pFrame = nullptr;
										pTangramPresentation->m_pTaskPanePage->CreateFrame(CComVariant(0), CComVariant(hWnd), CComBSTR(L"TaskPane"), &pFrame);
										pTangramPresentation->m_pTaskPaneFrame = (CWndFrame*)pFrame;
										if (pTangramPresentation->m_pTaskPaneFrame)
										{
											IWndNode* pNode = nullptr;
											pTangramPresentation->m_pTaskPaneFrame->Extend(CComBSTR("Default"), strXml.AllocSysString(), &pNode);
										}
									}
								}
								else
									pTangramPresentation->m_pTaskPaneFrame->ModifyHost(hWnd);
							}
						}
					}
				}
			}
			break;
			case 102:
			{
			}
			break;
			}

			return S_OK;
		}

		HRESULT CPowerPntCloudAddin::OnConnection(IDispatch* pHostApp, int ConnectMode)
		{
			if (::GetModuleHandle(_T("kso.dll")))
				return S_OK;
			if (m_strTemplateXml == _T(""))
			{
				CTangramXmlParse m_Parse;
				if (m_Parse.LoadXml(theApp.m_strConfigFile))
				{
					m_strRibbonXml = m_Parse[_T("RibbonUI")][_T("customUI")].xml();
					m_strDefaultTemplateXml = m_Parse[_T("DefaultTemplate")][_T("Tangrams")].xml();
				}
			}
			pHostApp->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pApplication);
			HRESULT hr = ((CPowerpointPlusEApplication*)this)->DispEventAdvise(m_pApplication);
			return S_OK;
		}

		HRESULT CPowerPntCloudAddin::OnDisconnection(int DisConnectMode)
		{
			HRESULT hr = ((CPowerpointPlusEApplication*)this)->DispEventUnadvise(m_pApplication);
			m_pApplication.Detach();
			return S_OK;
		}

		STDMETHODIMP CPowerPntCloudAddin::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml)
		{
			if (!RibbonXml)
				return E_POINTER;
			*RibbonXml = m_strRibbonXml.AllocSysString();
			return (*RibbonXml ? S_OK : E_OUTOFMEMORY);
		}

		bool CPowerPntCloudAddin::OnActiveOfficeObj(HWND hWnd)
		{
			auto it = m_mapOfficeObjects2.find(hWnd);
			if (it != m_mapOfficeObjects2.end())
			{
				if (m_pActivePowerPntObject != it->second)
				{
					m_pActivePowerPntObject = (CPowerPntObject*)it->second;
					OnPresentationActivate(m_pActivePowerPntObject);
				}
				return true;
			}
			return false;
		}

		void CPowerPntCloudAddin::OnPresentationActivate(CPowerPntObject* pPowerPntWnd)
		{
			CPowerPntObject* pWnd = pPowerPntWnd;
			CCloudAddinPresentation* pTangramPresentation = pPowerPntWnd->m_pPresentation;

			if (pTangramPresentation->m_pFrame)
			{
				pTangramPresentation->m_pFrame->ModifyHost((long)pWnd->m_hClient);
			}

			if (pTangramPresentation->m_pTaskPaneFrame)
			{
				if (::IsWindow(pWnd->m_hTaskPane))
				{
					pTangramPresentation->m_pTaskPaneFrame->ModifyHost((long)pWnd->m_hTaskPane);
				}
				else
					pTangramPresentation->m_pTaskPaneFrame->ModifyHost((long)pWnd->m_hTaskPaneChildWnd);
			}
			CWndFrame* pFrame = pPowerPntWnd->m_pPresentation->m_pFrame;
			if (pFrame)
			{
				if (pFrame->m_bDesignerState&&theApp.m_pDesignerFrame&&theApp.m_pDesigningFrame != pFrame)
				{
					theApp.m_pDesigningFrame = pFrame;
					pFrame->UpdateDesignerTreeInfo();
				}
			}
		}

		void CPowerPntCloudAddin::WindowDestroy(HWND hWnd)
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
			else if (strClassName == _T("mdiClass"))
			{
				OnCloseOfficeObj(strClassName, hWnd);
			}
		}

		void CPowerPntCloudAddin::WindowCreated(LPCTSTR lpszClass, LPCTSTR strName, HWND hPWnd, HWND hWnd)
		{
			CString strClassName = lpszClass;
			if (strClassName == _T("mdiClass"))
			{
				m_pActivePowerPntObject = new CComObject<CPowerPntObject>;
				m_pActivePowerPntObject->m_hForm = ::GetParent(hPWnd);
				m_pActivePowerPntObject->m_hClient = hPWnd;
				m_pActivePowerPntObject->m_hChildClient = hWnd;
				m_mapOfficeObjects2[m_pActivePowerPntObject->m_hForm] = m_pActivePowerPntObject;
				m_mapOfficeObjects[hWnd] = m_pActivePowerPntObject;
				::PostMessage(m_hHostWnd, WM_OFFICEOBJECTCREATED, (WPARAM)hWnd, 0);
			}
		}

		void CPowerPntCloudAddin::ConnectOfficeObj(HWND hWnd)
		{
			if (m_pApplication == nullptr)
				return;
			auto it = m_mapOfficeObjects.find(hWnd);
			CPowerPntObject* pPowerPntWndObj = (CPowerPntObject*)it->second;
			CComPtr<_Presentation> pDoc;
			m_pApplication->get_ActivePresentation(&pDoc);
			if (pDoc)
			{
				CComBSTR bstrName(L"");
				pDoc->get_Path(&bstrName);
				CString strPath = OLE2T(bstrName);
				pDoc->get_FullName(&bstrName);
				CString strKey = OLE2T(bstrName);

				bool bNewWindow = false;
				CCloudAddinPresentation* pTangramPresentation = nullptr;// 
				auto it = m_mapTangramPres.find(strKey);
				if (it != m_mapTangramPres.end())
				{
					pTangramPresentation = it->second;
					bNewWindow = true;
				}
				else
				{
					pTangramPresentation = new CCloudAddinPresentation();
					m_mapTangramPres[strKey] = pTangramPresentation;
				}
				pTangramPresentation->m_strKey = strKey;
				(*pTangramPresentation)[hWnd] = pPowerPntWndObj;
				pPowerPntWndObj->m_pPresentation = pTangramPresentation;
				if (bNewWindow)
					return;

				if (strPath == _T(""))
				{
					CString strTemplate = GetDocTemplateXml(_T("Please select Slild Template:"), theApp.m_strExeName);
					CTangramXmlParse m_Parse;
					bool bLoad = m_Parse.LoadXml(strTemplate);
					if (bLoad == false)
						bLoad = m_Parse.LoadFile(strTemplate);
					if (bLoad == false)
						return;
					VARIANT_BOOL bRet = false;
					CString strNewDocInfo = _T("");
					pTangramPresentation->m_strTaskPaneTitle = m_Parse.attr(_T("Title"), _T("TaskPane"));
					pTangramPresentation->m_nWidth = m_Parse.attrInt(_T("Width"), 200);
					pTangramPresentation->m_nHeight = m_Parse.attrInt(_T("Height"), 300);
					pTangramPresentation->m_nMsoCTPDockPosition = (MsoCTPDockPosition)m_Parse.attrInt(_T("DockPosition"), 4);
					pTangramPresentation->m_nMsoCTPDockPositionRestrict = (MsoCTPDockPositionRestrict)m_Parse.attrInt(_T("DockPositionRestrict"), 3);
					pPowerPntWndObj->m_pPresentation->m_strDocXml = m_Parse.xml();
					AddDocXml(pDoc, CComBSTR(pPowerPntWndObj->m_pPresentation->m_strDocXml), CComBSTR(L"Tangrams"), &bRet);
					CTangramXmlParse* pParse = m_Parse.GetChild(_T("TaskPaneUI"));
					if (pParse)
					{
						CString strXml = pParse->xml();
						if (strXml != _T(""))
						{
							pPowerPntWndObj->m_pPresentation->m_strTaskPaneXml = strXml;
						}
					}
					pParse = m_Parse.GetChild(_T("DocumentUI"));
					if (pParse)
					{
						IWndPage* pPage = nullptr;
						CreateWndPage((long)pPowerPntWndObj->m_hForm, &pPage);
						pPowerPntWndObj->m_pPresentation->m_pPage = (CWndPage*)pPage;
						IWndFrame* pFrame = nullptr;
						pPage->CreateFrame(CComVariant(0), CComVariant((long)pPowerPntWndObj->m_hClient), CComBSTR(L"Document"), &pFrame);
						pPowerPntWndObj->m_pPresentation->m_pFrame = (CWndFrame*)pFrame;
						IWndNode* pNode = nullptr;
						pFrame->Extend(CComBSTR(L"Default"), CComBSTR(pParse->xml()), &pNode);
					}
				}
				else
				{
					CComBSTR bstrXml(L"");
					GetDocXmlByKey(pDoc, CComBSTR(L"Tangrams"), &bstrXml);
					CString strXML = OLE2T(bstrXml);
					pPowerPntWndObj->m_pPresentation->m_strDocXml = strXML;

					if (strXML != _T(""))
					{
						CTangramXmlParse m_Parse;
						if (m_Parse.LoadXml(strXML) || m_Parse.LoadFile(strXML))
						{
							CTangramXmlParse* pParse = m_Parse.GetChild(_T("TaskPaneUI"));
							if (pParse)
							{
								CString strXml = pParse->xml();
								if (pPowerPntWndObj != nullptr)
								{
									pPowerPntWndObj->m_pPresentation->m_strTaskPaneXml = strXml;
								}
							}
							pParse = m_Parse.GetChild(_T("DocumentUI"));
							if (pParse)
							{
								IWndPage* pPage = nullptr;
								CreateWndPage((long)pPowerPntWndObj->m_hForm, &pPage);
								pPowerPntWndObj->m_pPresentation->m_pPage = (CWndPage*)pPage;
								IWndFrame* pFrame = nullptr;
								pPage->CreateFrame(CComVariant(0), CComVariant((long)pPowerPntWndObj->m_hClient), CComBSTR(L"Document"), &pFrame);
								pPowerPntWndObj->m_pPresentation->m_pFrame = (CWndFrame*)pFrame;
								IWndNode* pNode = nullptr;
								pFrame->Extend(CComBSTR(L"Default"), CComBSTR(pParse->xml()), &pNode);
							}
						}
					}
				}
			}
		}
	}
}// namespace OfficeCloudPlus

