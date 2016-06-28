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
#include "../../NodeWnd.h"
#include "../../WndNode.h"
#include "../../WndFrame.h"
#include "../../TangramHtmlTreeWnd.h"
#include "fm20.h"
#include "WordAddin.h"
#include "WordPlusDoc.h"

namespace OfficeCloudPlus
{
	namespace WordPlus
	{
		CWordCloudAddin::CWordCloudAddin() :CCloudAddin()
		{
			m_pActiveWordObject = nullptr;
		}

		STDMETHODIMP CWordCloudAddin::AddDocXml(IDispatch* pDocdisp, BSTR bstrXml, BSTR bstrKey, VARIANT_BOOL* bSuccess)
		{
			CComQIPtr<Word::_Document> pDoc(pDocdisp);
			if (pDoc)
			{
				CComPtr<Office::_CustomXMLParts> pCustomXMLParts;
				pDoc->get_CustomXMLParts(&pCustomXMLParts);
				_AddDocXml(pCustomXMLParts.p, bstrXml, bstrKey, bSuccess);
			}

			return S_OK;
		}

		STDMETHODIMP CWordCloudAddin::GetDocXmlByKey(IDispatch* pDocdisp, BSTR bstrKey, BSTR* pbstrXml)
		{
			CComQIPtr<Word::_Document> pDoc(pDocdisp);
			if (pDoc)
			{
				CComPtr<Office::_CustomXMLParts> pCustomXMLParts;
				pDoc->get_CustomXMLParts(&pCustomXMLParts);
				if (pCustomXMLParts)
					_GetDocXmlByKey(pCustomXMLParts.p, bstrKey, pbstrXml);
			}
			return S_OK;
		}

		void CWordCloudAddin::OnOpenDoc(WPARAM wParam)
		{
			CWordDocument* pWordDoc = (CWordDocument*)wParam;
			CWordObject* pWordWnd = pWordDoc->begin()->second;
			CreateWndPage((long)pWordWnd->m_hChildClient, &pWordDoc->m_pPage);
			if (pWordDoc->m_pPage)
			{
				pWordDoc->m_pPage->put_External(pWordDoc->m_pDoc);
				IWndFrame* pFrame = NULL;
				pWordDoc->m_pPage->CreateFrame(CComVariant(0), CComVariant((long)pWordWnd->m_hChildClient), CComBSTR(L"Document"), &pFrame);
				pWordDoc->m_pFrame = (CWndFrame*)pFrame;
				if (pWordDoc->m_pFrame)
				{
					auto it = pWordDoc->m_mapDocUIInfo.find(_T("Document"));
					CString strXml = _T("");
					if (it != pWordDoc->m_mapDocUIInfo.end())
						strXml = it->second;
					else
						strXml = _T("<Document><window><node name=\"Start\" id=\"HostView\"/></window></Document>");
					IWndNode* pNode = nullptr;
					pWordDoc->m_pFrame->Extend(CComBSTR(L"Default"), strXml.AllocSysString(), &pNode);
					CWndNode* _pNode = (CWndNode*)pNode;
					if (_pNode->m_pOfficeObj == nullptr)
					{
						_pNode->m_pOfficeObj = pWordDoc->m_pDoc;
						_pNode->m_pOfficeObj->AddRef();
					}
				}
			}
		}

		void CWordCloudAddin::OnDocumentBeforeSave(_Document* Doc, VARIANT_BOOL* SaveAsUI, VARIANT_BOOL* Cancel)
		{
			auto it = find(Doc);
			if (it != end())
			{
				CWndFrame* pFrame = it->second->m_pFrame;
				if (pFrame)
				{
					pFrame->UpdateWndNode();
				}
				pFrame = it->second->m_pTaskPaneFrame;
				if (pFrame)
				{
					pFrame->UpdateWndNode();
				}
			}
		}

		void CWordCloudAddin::WindowDestroy(HWND hWnd)
		{
			::GetClassName(hWnd, theApp.m_szBuffer, MAX_PATH);
			CString strClassName = CString(theApp.m_szBuffer);
			if (strClassName == _T("ThunderDFrame"))
			{
				auto it = m_mapVBAForm.find(hWnd);
				if (it != m_mapVBAForm.end())
					m_mapVBAForm.erase(it);
			}
			else if (strClassName == _T("_WwB"))
			{
				OnCloseOfficeObj(strClassName, hWnd);
			}
		}

		void CWordCloudAddin::WindowCreated(LPCTSTR lpszClass, LPCTSTR strName, HWND hPWnd, HWND hWnd)
		{
			CString strClassName = lpszClass;
			if (strClassName == _T("_WwB"))
			{
				::PostMessage(m_hHostWnd, WM_OFFICEOBJECTCREATED, (WPARAM)hWnd, 0);
			}
		}

		STDMETHODIMP CWordCloudAddin::TangramCommand(IDispatch* RibbonControl)
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

			CWordObject* pWnd = m_pActiveWordObject;
			CWordDocument* pWordDoc = m_pActiveWordObject->m_pWordPlusDoc;

			switch (nCmdIndex)
			{
			case 100:
			{
				if (pWnd)
				{
					CWndFrame* pFrame = pWnd->m_pWordPlusDoc->m_pFrame;
					if (pWnd->m_bDesignState == false)
					{
						pFrame->m_bDesignerState = true;
						pWnd->m_bDesignState = true;
						CreateCommonDesignerToolBar();
						CWndNode* pNode = pFrame->m_pWorkNode;
						if (pNode->m_strID.CompareNoCase(_T("hostView")) == 0)
						{
							CString strXml = _T("<Document><window><node name=\"Start\" /></window></Document>");
							IWndNode* pDesignNode = nullptr;
							pFrame->Extend(CComBSTR(L"default-inDesigning"), CComBSTR(strXml), &pDesignNode);
						}

						theApp.m_pDesigningFrame = pFrame;
						theApp.m_pDesigningFrame->UpdateDesignerTreeInfo();
						return 0;
					}
					else
					{
						pFrame->m_bDesignerState = false;
						pWnd->m_bDesignState = false;
					}
				}
			}
			break;
			case 101:
			case 102:
			{
				auto iter = m_mapTaskPaneMap.find((long)m_pActiveWordObject->m_hForm);
				if (iter != m_mapTaskPaneMap.end())
					m_mapTaskPaneMap[(long)m_pActiveWordObject->m_hForm]->put_Visible(true);
				else
				{
					CString strKey = _T("TaskPaneUI");
					auto it = pWnd->m_pWordPlusDoc->m_mapDocUIInfo.find(strKey);
					CString strXml = _T("");
					if (it != pWnd->m_pWordPlusDoc->m_mapDocUIInfo.end())
						strXml = it->second;
					else
						strXml = _T("<TaskPaneUI><window><node name=\"Start\" /></window></TaskPaneUI>");
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
							CTangramXmlParse* pNodeParse = m_Parse.FindItem(_T("node"));
							strCap = pNodeParse->attr(_T("caption"), _T(""));
							if (strCap == _T(""))
								strCap = pNodeParse->attr(_T("name"), _T(""));;
							CComBSTR bstrCap(strCap);
							HRESULT hr = m_pCTPFactory->CreateCTP(L"Tangram.Ctrl.1", bstrCap, vWindow, &_pCustomTaskPane);
							_pCustomTaskPane->AddRef();
							_pCustomTaskPane->put_Visible(true);
							m_mapTaskPaneMap[(long)m_pActiveWordObject->m_hForm] = _pCustomTaskPane;
							CComPtr<ITangramCtrl> pCtrlDisp;
							_pCustomTaskPane->get_ContentControl((IDispatch**)&pCtrlDisp);
							if (pCtrlDisp)
							{
								long hWnd = 0;
								pCtrlDisp->get_HWND(&hWnd);
								HWND hPWnd = ::GetParent((HWND)hWnd);
								pWnd->m_hTaskPane = (HWND)hWnd;
								CWindow m_Wnd;
								m_Wnd.Attach(hPWnd);
								m_Wnd.ModifyStyle(0, WS_CLIPSIBLINGS | WS_CLIPCHILDREN);
								m_Wnd.Detach();
								CWordDocument* m_pDoc = pWnd->m_pWordPlusDoc;
								if (m_pDoc->m_pTaskPanePage == nullptr)
								{
									HRESULT hr = theApp.m_pHostCore->CreateWndPage((long)hPWnd, &m_pDoc->m_pTaskPanePage);
									if (m_pDoc->m_pTaskPanePage)
									{
										IWndFrame* pTaskPaneFrame = nullptr;
										m_pDoc->m_pTaskPanePage->CreateFrame(CComVariant(0), CComVariant((long)hWnd), CComBSTR(L"TaskPane"), &pTaskPaneFrame);
										m_pDoc->m_pTaskPaneFrame = (CWndFrame*)pTaskPaneFrame;
										if (m_pDoc->m_pTaskPaneFrame)
										{
											IWndNode* pNode = nullptr;
											m_pDoc->m_pTaskPaneFrame->Extend(CComBSTR("Default"), strXml.AllocSysString(), &pNode);
											CWndNode* _pNode = (CWndNode*)pNode;
											if (_pNode->m_pOfficeObj == nullptr)
											{
												_pNode->m_pOfficeObj = pWordDoc->m_pDoc;
												_pNode->m_pOfficeObj->AddRef();
											}
										}
									}
								}
								else
									m_pDoc->m_pTaskPaneFrame->ModifyHost(hWnd);
							}
						}
					}
				}
				if (nCmdIndex == 102 && pWnd)
				{
					CreateCommonDesignerToolBar();
					CWndFrame* pFrame = pWnd->m_pWordPlusDoc->m_pTaskPaneFrame;
					if (pWnd->m_bDesignTaskPane == false)
					{
						pFrame->m_bDesignerState = true;
						if (theApp.m_pDesigningFrame != pFrame)
						{
							theApp.m_pDesigningFrame = pFrame;
							pFrame->UpdateDesignerTreeInfo();
						}
					}
					else
					{
						pFrame->m_bDesignerState = false;
						if (theApp.m_pDesigningFrame == pFrame)
						{
							theApp.m_pDesigningFrame = nullptr;
							pFrame->UpdateDesignerTreeInfo();
						}
					}
				}
			}
			break;
			}

			return S_OK;
		}

		HRESULT CWordCloudAddin::OnConnection(IDispatch* pHostApp, int ConnectMode)
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
			HRESULT hr = ((CWordAppEvents2*)this)->DispEventAdvise(m_pApplication);
			//hr = ((CWordAppEvents3*)this)->DispEventAdvise(m_pApplication);
			//hr = ((CWordAppEvents4*)this)->DispEventAdvise(m_pApplication);
			if (::IsWindow(m_hHostWnd))
			{
				WNDCLASSEX wcex;
				wcex.cbSize = sizeof(WNDCLASSEX);
				wcex.style = CS_HREDRAW | CS_VREDRAW;
				wcex.lpfnWndProc = ::DefWindowProc;
				wcex.cbClsExtra = 0;
				wcex.cbWndExtra = 0;
				wcex.hInstance = theApp.m_hInstance;
				wcex.hIcon = NULL;
				wcex.hCursor = NULL;
				wcex.hbrBackground = NULL;
				wcex.lpszMenuName = NULL;
				wcex.lpszClassName = L"Tangram Word Helper Window";
				wcex.hIconSm = NULL;
				RegisterClassEx(&wcex);
			}
			return S_OK;
		}

		HRESULT CWordCloudAddin::OnDisconnection(int DisConnectMode)
		{
			HRESULT hr = ((CWordAppEvents2*)this)->DispEventUnadvise(m_pApplication);
			m_pApplication = nullptr;
			return S_OK;
		}

		STDMETHODIMP CWordCloudAddin::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml)
		{
			if (!RibbonXml || m_strRibbonXml == _T(""))
				return E_POINTER;
			*RibbonXml = CComBSTR(m_strRibbonXml);
			return (*RibbonXml ? S_OK : E_OUTOFMEMORY);
		}

		CString CWordCloudAddin::GetFormXml(CString strFormName)
		{
			CWordDocument* pWordDoc = m_pActiveWordObject->m_pWordPlusDoc;
			auto it = pWordDoc->m_mapUserFormScript.find(strFormName);
			if (it != pWordDoc->m_mapUserFormScript.end())
				return it->second;

			return _T("");
		}

		STDMETHODIMP CWordCloudAddin::ExportOfficeObjXml(IDispatch* OfficeObject, BSTR* bstrXml)
		{
			return S_OK;
		}

		void CWordCloudAddin::UpdateOfficeObj(IDispatch* pObj, CString strXml, CString _strName)
		{
			CComQIPtr<MSForm::_UserForm> pForm(pObj);
			if (pForm)
			{
				CComPtr<_Document> pDoc;
				m_pApplication->get_ActiveDocument(&pDoc);
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
			CComQIPtr<_Document> pDoc(pObj);
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

		STDMETHODIMP CWordCloudAddin::CreateOfficeDocument(BSTR bstrXml)
		{
			CComPtr<Documents> pDocsDisp2;
			m_pApplication->get_Documents(&pDocsDisp2);
			if (pDocsDisp2)
			{
				theApp.Lock();
				auto it = theApp.m_mapValInfo.find(_T("doctemplate"));
				if (it != theApp.m_mapValInfo.end())
				{
					::VariantClear(&it->second);
					theApp.m_mapValInfo.erase(it);
				}
				CString strXml = OLE2T(bstrXml);
				if (strXml != _T(""))
					theApp.m_mapValInfo[_T("doctemplate")] = CComVariant(strXml);
				theApp.Unlock();

				CComPtr<_Document> pDoc;
				CComVariant varVisible(true);
				CComVariant varNew(true);
				CComVariant varTemplate(L"");
				CComVariant varTypr(0);
				pDocsDisp2->Add(&varTemplate, &varNew, &varTypr, &varVisible, &pDoc);
			}

			return S_OK;
		}

		bool CWordCloudAddin::OnActiveOfficeObj(HWND hWnd)
		{
			auto it = m_mapOfficeObjects2.find(hWnd);
			if (it != m_mapOfficeObjects2.end())
			{
				if (m_pActiveWordObject != it->second)
				{
					m_pActiveWordObject = (CWordObject*)it->second;
				}
				OnDocActivate(m_pActiveWordObject);
				return true;
			}
			return false;
		}

		void CWordCloudAddin::OnDocActivate(CWordObject* pWnd)
		{
			if (m_pActiveWordObject)
			{
				if (m_pActiveWordObject->m_bDesignState)
				{
					CreateCommonDesignerToolBar();
				}
				CWordDocument* pWordPlusDoc = m_pActiveWordObject->m_pWordPlusDoc;
				if (pWordPlusDoc)
				{
					if (pWordPlusDoc->m_pFrame)
						pWordPlusDoc->m_pFrame->ModifyHost((long)m_pActiveWordObject->m_hChildClient);
					if (pWordPlusDoc->m_pTaskPaneFrame)
					{
						if (m_pActiveWordObject->m_hTaskPane)
							pWordPlusDoc->m_pTaskPaneFrame->ModifyHost((long)m_pActiveWordObject->m_hTaskPane);
						else
							pWordPlusDoc->m_pTaskPaneFrame->ModifyHost((long)m_pActiveWordObject->m_hTaskPaneChildWnd);
					}
				}
			}
		}

		void CWordCloudAddin::ConnectOfficeObj(HWND hWnd)
		{
			if (m_pApplication == nullptr)
				return;
			::GetWindowText(hWnd, theApp.m_szBuffer, 255);
			CString strCaption = CString(theApp.m_szBuffer);
			if (strCaption == _T(""))
				return;

			m_pActiveWordObject = new CComObject<CWordObject>;
			m_pActiveWordObject->m_hClient = ::GetParent(hWnd);
			m_pActiveWordObject->m_hChildClient = hWnd;
			m_pActiveWordObject->m_hForm = ::GetParent(m_pActiveWordObject->m_hClient);
			m_mapOfficeObjects2[m_pActiveWordObject->m_hForm] = m_pActiveWordObject;
			m_mapOfficeObjects[hWnd] = m_pActiveWordObject;

			CWindow m_wnd;
			m_wnd.Attach(m_pActiveWordObject->m_hChildClient);
			m_wnd.ModifyStyle(WS_CAPTION | WS_SYSMENU | WS_MAXIMIZEBOX | WS_MINIMIZEBOX, 0);

			bool bNewWindow = false;
			_Document* pDoc = nullptr;
			m_pApplication->get_ActiveDocument(&pDoc);
			CWordDocument* pWordDoc = nullptr;
			auto it1 = find(pDoc);
			if (it1 != end())
			{
				pWordDoc = it1->second;
				bNewWindow = true;
			}
			else
			{
				pWordDoc = new CWordDocument();
				(*this)[pDoc] = pWordDoc;
				HRESULT hr = pWordDoc->DispEventAdvise(pDoc);
				pWordDoc->m_pDoc = pDoc;
				pWordDoc->m_pDoc->AddRef();
			}
			m_pActiveWordObject->m_pOfficeObj = pDoc;
			m_pActiveWordObject->m_pOfficeObj->AddRef();
			m_pActiveWordObject->m_pWordPlusDoc = pWordDoc;
			(*pWordDoc)[hWnd] = m_pActiveWordObject;
			if (bNewWindow)
				return;

			RECT rc;
			::GetClientRect(m_pActiveWordObject->m_hClient, &rc);
			::SetWindowPos(m_pActiveWordObject->m_hChildClient, HWND_TOP, 0, 0, rc.right, rc.bottom, SWP_NOREDRAW | SWP_NOACTIVATE);
			CComBSTR bstrPath(L"");
			pDoc->get_Path(&bstrPath);
			CString strPath = OLE2T(bstrPath);
			if (strPath == _T(""))
			{
				CString strTemplate = GetDocTemplateXml(_T("Please select Word Document Template:"), theApp.m_strExeName);

				CTangramXmlParse m_Parse;
				bool bLoad = m_Parse.LoadXml(strTemplate);
				if (bLoad == false)
					bLoad = m_Parse.LoadFile(strTemplate);
				if (bLoad == false)
					return;
				VARIANT_BOOL bRet = false;
				CComBSTR bstrRet(L"");
				CString strNewDocInfo = _T("");
				pWordDoc->m_strTaskPaneTitle = m_Parse.attr(_T("Title"), _T("TaskPane"));
				pWordDoc->m_nWidth = m_Parse.attrInt(_T("Width"), 200);
				pWordDoc->m_nHeight = m_Parse.attrInt(_T("Height"), 300);
				pWordDoc->m_nMsoCTPDockPosition = (MsoCTPDockPosition)m_Parse.attrInt(_T("DockPosition"), 4);
				pWordDoc->m_nMsoCTPDockPositionRestrict = (MsoCTPDockPositionRestrict)m_Parse.attrInt(_T("DockPositionRestrict"), 3);
				pWordDoc->m_strDocXml = m_Parse.xml();
			}
			else
			{
				auto it = m_mapTaskPaneMap.find((LONG)m_pActiveWordObject->m_hForm);
				if (it != m_mapTaskPaneMap.end())
				{
					it->second->put_Visible(false);
					m_mapTaskPaneMap.erase(it);
				}
			}
			m_pActiveWordObject->m_pWordPlusDoc->InitDocument();
		}
	}
}
