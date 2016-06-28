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
#include "../../NodeWnd.h"
#include "../../WndNode.h"
#include "../../WndFrame.h"
#include "../../TangramHtmlTreeWnd.h"
#include "fm20.h"
#include "ExcelAddin.h"
#include "ExcelPlusWnd.h"
#include "ExcelPlusEvents.cpp"
//#include "../../EclipsePlus\EclipseMain.h"

namespace OfficeCloudPlus
{
	namespace ExcelPlus
	{
		CExcelCloudAddin::CExcelCloudAddin():CCloudAddin()
		{
			m_pActiveExcelObject = nullptr;
			m_bOldVer = false;
			CString strVer = theApp.GetFileVer();
			int nPos = strVer.Find(_T("."));
			strVer = strVer.Left(nPos);
			int nVer = _wtoi(strVer);
			if (nVer<15)
				m_bOldVer = true;
		}

		CExcelCloudAddin::~CExcelCloudAddin()
		{
		}

		STDMETHODIMP CExcelCloudAddin::AddDocXml(IDispatch* pDocdisp, BSTR bstrXml, BSTR bstrKey, VARIANT_BOOL* bSuccess)
		{
			CComQIPtr<Excel::_Workbook> pDoc(pDocdisp);
			if (pDoc)
			{
				CComPtr<Office::_CustomXMLParts> pCustomXMLParts;
				pDoc->get_CustomXMLParts(&pCustomXMLParts);
				_AddDocXml(pCustomXMLParts.p, bstrXml, bstrKey, bSuccess);
			}

			return S_OK;
		}

		STDMETHODIMP CExcelCloudAddin::GetDocXmlByKey(IDispatch* pDocdisp, BSTR bstrKey, BSTR* pbstrXml)
		{
			CComQIPtr<Excel::_Workbook> pDoc(pDocdisp);
			if (pDoc)
			{
				CComPtr<Office::_CustomXMLParts> pCustomXMLParts;
				pDoc->get_CustomXMLParts(&pCustomXMLParts);
				_GetDocXmlByKey(pCustomXMLParts.p, bstrKey, pbstrXml);
			}
			return S_OK;
		}

		void CExcelCloudAddin::OnWorkBookActivate(CExcelObject* pExcelObject)
		{
			CExcelWorkBook* pWorkBook = pExcelObject->m_pWorkBook;
			if (pWorkBook == nullptr)
				return;
			HWND hWnd = (HWND)pExcelObject->m_hForm;
			if (pWorkBook->m_pFrame)
				pWorkBook->m_pFrame->ModifyHost((long)pExcelObject->m_hChildClient);
			if (pWorkBook->m_pTaskPaneFrame)
			{
				if (pExcelObject->m_hTaskPane)
				{
					pWorkBook->m_pTaskPaneFrame->ModifyHost((long)pExcelObject->m_hTaskPane);
				}
				else 
					pWorkBook->m_pTaskPaneFrame->ModifyHost((long)pExcelObject->m_hTaskPaneChildWnd);
			}

			if (pExcelObject->m_strActiveSheetName != _T(""))
			{
				CString strName = pExcelObject->m_strActiveSheetName;
				auto it = pWorkBook->m_mapWorkSheetInfo.find(strName);
				CString strXml = _T("");
				if (it != pWorkBook->m_mapWorkSheetInfo.end())
				{
					strXml = it->second;
				}
				else
				{
					strXml = pWorkBook->m_mapWorkSheetInfo[_T("Default")];
				}
					
				IWndNode* pNode = nullptr;
				pWorkBook->m_pFrame->Extend(strName.AllocSysString(), CComBSTR(strXml), &pNode);
			}

			CWndFrame* pFrame = pExcelObject->m_pWorkBook->m_pFrame;
			if (pFrame)
			{
				if (pFrame->m_bDesignerState&&theApp.m_pDesignerFrame&&theApp.m_pDesigningFrame != pFrame)
				{
					theApp.m_pDesigningFrame = pFrame;
					pFrame->UpdateDesignerTreeInfo();
				}
			}
		}

		CString CExcelCloudAddin::GetFormXml(CString strFormName)
		{
			CExcelWorkBook* pWorkBook = m_pActiveExcelObject->m_pWorkBook;
			auto it = pWorkBook->m_mapUserFormScript.find(strFormName);
			if (it != pWorkBook->m_mapUserFormScript.end())
				return it->second;
			return _T("");
		}

		void CExcelCloudAddin::OnOpenDoc(WPARAM wParam) 
		{
			CExcelObject* pObj = (CExcelObject*)wParam;
			if (pObj)
			{
				pObj->OnOpenDoc();
			}
		}
		
		void CExcelCloudAddin::SetFocus(HWND hWnd)
		{
			if (m_pActiveExcelObject == nullptr)
				return;
			if (hWnd == m_pActiveExcelObject->m_hExcelEdit)
			{
				theApp.m_pWndNode = nullptr;
				if (::IsWindowVisible(m_pActiveExcelObject->m_hChildClient) == FALSE)
				{
					::PostMessage(m_pActiveExcelObject->m_hChildClient, WM_SETWNDFOCUSE, 0, 0);
				}
			}
		}

		void CExcelCloudAddin::ProcessMsg(LPMSG lpMsg)
		{
			if (m_pActiveExcelObject == nullptr)
				return;
			switch (lpMsg->message)
			{
			case WM_LBUTTONDOWN:
			case WM_RBUTTONDOWN:
				m_pActiveExcelObject->ProcessMouseDownMsg(lpMsg->hwnd);
				break;
			case WM_SETWNDFOCUSE:
				if (lpMsg->wParam)
				{
					if (::IsWindowEnabled(m_pActiveExcelObject->m_hExcelEdit))
						::EnableWindow(m_pActiveExcelObject->m_hExcelEdit, false);
					::SendMessage((HWND)lpMsg->wParam, WM_MOUSEACTIVATE, (WPARAM)::GetActiveWindow(), 0);
				}
				else
				{
					::ShowWindow(m_pActiveExcelObject->m_hExcelEdit2, SW_HIDE);
				}

				break;
			}
		}

		void CExcelCloudAddin::OnVbaFormInit(HWND hWnd, CWndFrame* pFrame)
		{
			::SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)m_pActiveExcelObject->m_pWorkBook);
		}

		void CExcelCloudAddin::UpdateOfficeObj(IDispatch* pObj, CString strXml, CString _strName)
		{
			CComQIPtr<MSForm::_UserForm> pForm(pObj);
			if (pForm)
			{
				CComPtr<Excel::_Workbook> pWb;
				m_pExcelApplication->get_ActiveWorkbook(&pWb);
				if (pWb)
				{
					CComPtr<Office::_CustomXMLParts> pCustomXMLParts;
					pWb->get_CustomXMLParts(&pCustomXMLParts);
					CComBSTR bstrXml(L"");
					_GetDocXmlByKey(pCustomXMLParts.p, CComBSTR(L"workbook"), &bstrXml);
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
								strXml2.Format(_T("<%s>%s</%s>"), strName,strXml,strName);
								CTangramXmlParse m_Parse2;
								m_Parse2.LoadXml(strXml2);
								pParse->AddNode(&m_Parse2, _T(""));
								strXml2 = m_Parse.xml();
							}
							VARIANT_BOOL bRet = false;
							AddDocXml(pWb, CComBSTR(strXml2), CComBSTR(L"workbook"), &bRet);
						}
					}
				}
				return;
			}
			CComQIPtr<Excel::_Worksheet> pSheet(pObj);
			if (pSheet)
			{
				CComPtr<Excel::CustomProperties> pCustomProperties;
				pSheet->get_CustomProperties(&pCustomProperties);
				Excel::ICustomProperties* pProps = (Excel::ICustomProperties*)pCustomProperties.p;
				long nCount = 0;
				CComBSTR bstrName(L"");
				if (pProps == nullptr)
					return;
				pProps->get_Count(&nCount);
				for (int i = 1; i <= nCount; i++)
				{
					CComPtr<Excel::CustomProperty> pProp;
					pProps->get_Item(CComVariant(i), &pProp);
					Excel::ICustomProperty* _pProp = (ICustomProperty*)pProp.p;
					_pProp->get_Name(&bstrName);
					CString strName = OLE2T(bstrName);
					if (strName == _T("Tangram"))
					{
						CComVariant var;
						_pProp->get_Value(&var);
						CString _strXml = OLE2T(var.bstrVal);
						::VariantClear(&var);
						CString strKey = _T("</") + _strName + _T(">");
						int nPos = _strXml.Find(strKey);
						CString str1 = _strXml.Mid(nPos);
						strKey = _T("<") + _strName +_T(" ");
						CString str2 = _T("");
						nPos = _strXml.Find(strKey);
						if (nPos == -1)
						{
							strKey = _T("<") + _strName + _T(">");
							nPos = _strXml.Find(strKey);
							str2 = _strXml.Left(nPos);
							str2 += strKey;
						}
						else
						{
							str2 = _strXml.Left(nPos);
							str2 += _T("<") + _strName + _T(">");
						}
						_strXml = str2 + strXml + str1;
						_pProp->put_Value(CComVariant(_strXml));
						return;
					}
				}
			}

			if (pObj)
			{
				CComBSTR szMember(L"Parent");
				DISPID dispid = -1;
				HRESULT hr = pObj->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_USER_DEFAULT, &dispid);
				if (hr == S_OK)
				{
					DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
					VARIANT result = { 0 };
					EXCEPINFO excepInfo;
					memset(&excepInfo, 0, sizeof excepInfo);
					UINT nArgErr = (UINT)-1;
					HRESULT hr = pObj->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &result, &excepInfo, &nArgErr);
					if (S_OK == hr && VT_DISPATCH == result.vt)
					{
						CComQIPtr<Excel::_Workbook> pWb(result.pdispVal);
						if (pWb)
						{
							CComPtr<Office::_CustomXMLParts> pCustomXMLParts;
							pWb->get_CustomXMLParts(&pCustomXMLParts);
							CComBSTR bstrXml(L"");
							_GetDocXmlByKey(pCustomXMLParts.p, CComBSTR(L"workbook"), &bstrXml);
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
											AddDocXml(pWb, CComBSTR(str), CComBSTR(L"workbook"), &bRet);
										}
									}
								}
							}
						}
					}
					::VariantClear(&result);
				}
			}

		}
		
		STDMETHODIMP CExcelCloudAddin::TangramCommand(IDispatch* RibbonControl)
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
			CExcelObject* pWnd = m_pActiveExcelObject;
			CExcelWorkBook* m_pWorkBook = pWnd->m_pWorkBook;
			switch (nCmdIndex)
			{
			case 100:
			{
				CWndFrame* pFrame = m_pWorkBook->m_pTaskPaneFrame;
				if (pFrame)
				{
					if (pFrame->m_bDesignerState == false)
					{
						pFrame->m_bDesignerState = true;
						if (theApp.m_pDesigningFrame != pFrame)
						{
							theApp.m_pDesigningFrame = pFrame;
							pFrame->UpdateDesignerTreeInfo();
						}
					}
				}

				if (m_pWorkBook->m_pFrame)
				{
					m_pWorkBook->m_pFrame->m_bDesignerState = true;
					CreateCommonDesignerToolBar();

					CWndNode* pNode = theApp.m_pDesignerFrame->m_pWorkNode;
					theApp.HostPosChanged(pNode);
					theApp.m_pDesigningFrame = m_pWorkBook->m_pFrame;
					theApp.m_pDesigningFrame->UpdateDesignerTreeInfo();
				}
			}
			break;
			case 101:
				//case 102:
			{
				auto iter = m_mapTaskPaneMap.find((long)m_pActiveExcelObject->m_hForm);
				if (iter != m_mapTaskPaneMap.end())
					m_mapTaskPaneMap[(long)m_pActiveExcelObject->m_hForm]->put_Visible(true);
				else
				{
					CExcelCloudAddin* pAddin = (CExcelCloudAddin*)theApp.m_pHostCore;
					IDispatch* pDisp = nullptr;
					m_pWorkBook->m_pWorkBook->get_ActiveSheet(&pDisp);
					CComQIPtr<Excel::_Worksheet> pSheet(pDisp);

					CString strName = _T("");
					CString strXml = m_pWorkBook->m_strDefaultSheetXml;
					if (pSheet)
					{
						CComPtr<Excel::CustomProperties> pCustomProperties;
						pSheet->get_CustomProperties(&pCustomProperties);
						Excel::ICustomProperties* pProps = (Excel::ICustomProperties*)pCustomProperties.p;
						long nCount = 0;
						CComBSTR bstrName(L"");
						pProps->get_Count(&nCount);
						for (int i = 1; i <= nCount; i++)
						{
							CComPtr<Excel::CustomProperty> pProp;
							pProps->get_Item(CComVariant(i), &pProp);
							Excel::ICustomProperty* _pProp = (ICustomProperty*)pProp.p;
							_pProp->get_Name(&bstrName);
							strName = OLE2T(bstrName);
							if (strName == _T("Tangram"))
							{
								CComVariant var;
								_pProp->get_Value(&var);
								strXml = OLE2T(var.bstrVal);
								::VariantClear(&var);
								break;
							}
						}
					}
					if (strXml == _T(""))
						return 0;
					CTangramXmlParse m_Parse;
					bool b = m_Parse.LoadXml(strXml);
					if (b)
					{
						CTangramXmlParse* pParse = m_Parse.GetChild(_T("default"));
						if (pParse)
						{
							pParse = pParse->GetChild(_T("taskpane"));
							if (pParse)
							{
								strName += _T("_taskpane");
								Office::_CustomTaskPane* _pCustomTaskPane = nullptr;
								CString strSheetName = _T("");
								auto it = pAddin->m_mapTaskPaneMap.find((long)m_pActiveExcelObject->m_hForm);
								if (it == pAddin->m_mapTaskPaneMap.end())
								{
									VARIANT vWindow;
									vWindow.vt = VT_DISPATCH;
									vWindow.pdispVal = nullptr;
									HRESULT hr = pAddin->m_pCTPFactory->CreateCTP(L"Tangram.Ctrl.1", m_pWorkBook->m_strTaskPaneTitle.AllocSysString(), vWindow, &_pCustomTaskPane);
									_pCustomTaskPane->AddRef();
									pAddin->m_mapTaskPaneMap[(long)m_pActiveExcelObject->m_hForm] = _pCustomTaskPane;
									CComPtr<ITangramCtrl> pCtrlDisp;
									_pCustomTaskPane->get_ContentControl((IDispatch**)&pCtrlDisp);
									if (pCtrlDisp)
									{
										long hWnd = 0;
										pCtrlDisp->get_HWND(&hWnd);
										pWnd->m_hTaskPane = (HWND)hWnd;
										if (m_pWorkBook->m_pTaskPanePage == nullptr)
										{
											HWND hPWnd = ::GetParent((HWND)hWnd);
											CWindow m_Wnd;
											m_Wnd.Attach(hPWnd);
											m_Wnd.ModifyStyle(0, WS_CLIPSIBLINGS | WS_CLIPCHILDREN);
											m_Wnd.Detach();
											HRESULT hr = theApp.m_pHostCore->CreateWndPage((long)hPWnd, &m_pWorkBook->m_pTaskPanePage);
											if (m_pWorkBook->m_pTaskPanePage)
											{
												IWndFrame* pTaskPaneFrame = nullptr;
												m_pWorkBook->m_pTaskPanePage->CreateFrame(CComVariant(0), CComVariant((long)hWnd), CComBSTR(L"TaskPane"), &pTaskPaneFrame);
												m_pWorkBook->m_pTaskPaneFrame = (CWndFrame*)pTaskPaneFrame;
												if (m_pWorkBook->m_pTaskPaneFrame)
												{
													IWndNode* pNode = nullptr;
													m_pWorkBook->m_pTaskPaneFrame->Extend(CComBSTR(strName), pParse->xml().AllocSysString(), &pNode);
													CWndNode* _pNode = (CWndNode*)pNode;
													if (_pNode->m_pOfficeObj == nullptr&&pDisp)
													{
														_pNode->m_pOfficeObj = pDisp;
														_pNode->m_pOfficeObj->AddRef();
													}
												}
											}
											_pCustomTaskPane->put_DockPosition(m_pWorkBook->m_nMsoCTPDockPosition);
											_pCustomTaskPane->put_DockPositionRestrict(m_pWorkBook->m_nMsoCTPDockPositionRestrict);
											_pCustomTaskPane->put_Width(m_pWorkBook->m_nWidth);
											_pCustomTaskPane->put_Height(m_pWorkBook->m_nHeight);
											_pCustomTaskPane->put_Visible(true);
										}
										else
											m_pWorkBook->m_pTaskPaneFrame->ModifyHost(hWnd);
									}
								}
								else
								{
									_pCustomTaskPane = it->second;
									_pCustomTaskPane->put_Visible(true);
									if (m_pWorkBook->m_pTaskPaneFrame)
									{
										IWndNode* pNode = nullptr;
										m_pWorkBook->m_pTaskPaneFrame->Extend(CComBSTR(strName), pParse->xml().AllocSysString(), &pNode);
									}
								}
							}
						}
					}
				}
				if (nCmdIndex == 102)
				{
					CWndFrame* pFrame = m_pWorkBook->m_pTaskPaneFrame;
					if (pFrame == nullptr)return S_OK;
					CreateCommonDesignerToolBar();
					if (pFrame->m_bDesignerState == false)
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
				//InitEclipse();
			}
			break;
			case 102:
			{
				CString strXml = _T("");
				if (strXml == _T(""))
				{
					CString str = theApp.m_strExeName;
					CString strCaption = _T("Excel Plus");
					str += _T("\\");
					strXml = GetDocTemplateXml(strCaption, str);
				}
			}
			break;
			}

			return S_OK;
		}

		HRESULT CExcelCloudAddin::OnConnection(IDispatch* pHostApp, int ConnectMode)
		{
			if (::GetModuleHandle(_T("kso.dll")))
				return S_OK;
			if (m_pExcelApplication == nullptr)
			{
				pHostApp->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pExcelApplication);
			}
			if (m_strTemplateXml == _T(""))
			{
				CTangramXmlParse m_Parse;
				if (m_Parse.LoadXml(theApp.m_strConfigFile))
				{
					m_strRibbonXml = m_Parse[_T("RibbonUI")][_T("customUI")].xml();
					m_strDefaultTemplateXml = m_Parse[_T("DefaultTemplate")][_T("workbook")].xml();
				}
			}

			return S_OK;
		}

		STDMETHODIMP CExcelCloudAddin::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml)
		{
			if (!RibbonXml)
				return E_POINTER;
			*RibbonXml = m_strRibbonXml.AllocSysString();
			return (*RibbonXml ? S_OK : E_OUTOFMEMORY);
		}

		void CExcelCloudAddin::WindowCreated(LPCTSTR lpszClass, LPCTSTR strName, HWND hPWnd, HWND hWnd)
		{
			CString strClassName = lpszClass;
			if (strClassName == _T("EXCEL7"))
			{
				m_pActiveExcelObject = new CComObject<CExcelObject>;
				m_pActiveExcelObject->m_hClient = hPWnd;
				m_pActiveExcelObject->m_hChildClient = hWnd;
				m_pActiveExcelObject->m_hForm = ::GetParent(hPWnd);
				m_pActiveExcelObject->m_hExcelEdit = ::FindWindowEx(::GetParent(hPWnd), NULL, _T("EXCEL<"), NULL);;
				m_pActiveExcelObject->m_hExcelEdit2 = ::FindWindowEx(hPWnd, NULL, _T("EXCEL6"), NULL);
				m_mapOfficeObjects[hWnd] = m_pActiveExcelObject;
				m_mapOfficeObjects2[m_pActiveExcelObject->m_hForm] = m_pActiveExcelObject;
				::PostMessage(m_hHostWnd, WM_OFFICEOBJECTCREATED, (WPARAM)hWnd, 0);
			}
		}

		void CExcelCloudAddin::WindowDestroy(HWND hWnd)
		{
			::GetClassName(hWnd, theApp.m_szBuffer, MAX_PATH);
			CString strClassName = CString(theApp.m_szBuffer);
			if (strClassName == _T("ThunderDFrame"))
			{
				CExcelWorkBook* pWorkBook = (CExcelWorkBook*)::GetWindowLongPtr(hWnd, GWLP_USERDATA);
				if (pWorkBook)
				{

				}
				auto it = m_mapVBAForm.find(hWnd);
				if (it != m_mapVBAForm.end())
					m_mapVBAForm.erase(it);
			}
			else if (strClassName == _T("EXCEL="))
			{
				m_pActiveExcelObject->SheetNameChanged();
			}
			else if (strClassName == _T("EXCEL7"))
			{
				OnCloseOfficeObj(strClassName, hWnd);
				return;
			}
		}

		STDMETHODIMP CExcelCloudAddin::ExportOfficeObjXml(IDispatch* OfficeObject, BSTR* bstrXml)
		{
			return S_OK;
		}

		STDMETHODIMP CExcelCloudAddin::CreateOfficeDocument(BSTR bstrXml)
		{
			CComPtr<Workbooks> pWorkBooks;
			m_pExcelApplication->get_Workbooks(&pWorkBooks);
			if (pWorkBooks)
			{
				theApp.Lock();
				auto it = theApp.m_mapValInfo.find(_T("doctemplate"));
				if (it != theApp.m_mapValInfo.end())
				{
					::VariantClear(&it->second);
					theApp.m_mapValInfo.erase(it);
				}
				CString strXml = OLE2T(bstrXml);
				if(strXml!=_T(""))
					theApp.m_mapValInfo[_T("doctemplate")] = CComVariant(strXml);
				theApp.Unlock();
				CComPtr<_Workbook> pDoc;
				CComVariant varTemplate(L"");
				pWorkBooks->Add(varTemplate, 0, &pDoc);
			}

			return S_OK;
		}

		bool CExcelCloudAddin::OnActiveOfficeObj(HWND hWnd)
		{
			auto it = m_mapOfficeObjects2.find(hWnd);
			if (it != m_mapOfficeObjects2.end())
			{
				if (m_pActiveExcelObject != it->second)
				{
					m_pActiveExcelObject = (CExcelObject*)it->second;
				}
				OnWorkBookActivate((CExcelObject*)it->second);
				return true;
			}
			return false;
		}

		void CExcelCloudAddin::ConnectOfficeObj(HWND hWnd)
		{
			if (m_pExcelApplication==nullptr)
				return;
			auto it = m_mapOfficeObjects.find(hWnd);
			CExcelObject* pExcelObj = (CExcelObject*)it->second;
			Excel::_Workbook* pWorkBook = nullptr;
			m_pExcelApplication->get_ActiveWorkbook(&pWorkBook);
			pExcelObj->m_pOfficeObj = pWorkBook;
			pExcelObj->m_pOfficeObj->AddRef();
			pExcelObj->m_hTaskPaneWnd = ::CreateWindowEx(NULL, _T("Tangram Remote Helper Window"), L"", WS_CHILD, 0, 0, 0, 0, pExcelObj->m_hChildClient, NULL, theApp.m_hInstance, NULL);
			pExcelObj->m_hTaskPaneChildWnd = ::CreateWindowEx(NULL, _T("Tangram Remote Helper Window"), L"", WS_CHILD, 0, 0, 0, 0, (HWND)pExcelObj->m_hTaskPaneWnd, NULL, theApp.m_hInstance, NULL);
			bool bNewWindow = false;
			CExcelWorkBook* _pWorkBook = nullptr;
			auto it1 = find(pWorkBook);
			if (it1 != end())
			{
				_pWorkBook = it1->second;
				bNewWindow = true;
			}
			else
			{
				_pWorkBook = new CExcelWorkBook();
				(*this)[pWorkBook] = _pWorkBook;
				HRESULT hr = _pWorkBook->DispEventAdvise(pWorkBook);
				_pWorkBook->m_pWorkBook = pWorkBook;
				_pWorkBook->m_pWorkBook->AddRef();
			}
			pExcelObj->m_pWorkBook = _pWorkBook;
			_pWorkBook->m_mapExcelWorkBookWnd[pExcelObj->m_hChildClient] = pExcelObj;
			if (bNewWindow)
				return;
			CComBSTR bstrPath(L"");
			pWorkBook->get_Path(0, &bstrPath);
			CString strPath = OLE2T(bstrPath);
			if (strPath == _T(""))
			{
				CString strTemplate = GetDocTemplateXml(_T("Please select WorkBook Template:"), theApp.m_strExeName);

				CTangramXmlParse m_Parse;
				bool bLoad = m_Parse.LoadXml(strTemplate);
				if (bLoad == false)
					bLoad = m_Parse.LoadFile(strTemplate);
				if (bLoad == false)
					return;

				CComPtr<Excel::Sheets> pSheets;
				pWorkBook->get_Sheets(&pSheets);
				CComPtr<IDispatch> pActive;
				pWorkBook->get_ActiveSheet(&pActive);

				_pWorkBook->m_strTaskPaneTitle = m_Parse.attr(_T("Title"), _T("TaskPane"));
				_pWorkBook->m_nWidth = m_Parse.attrInt(_T("Width"), 200);
				_pWorkBook->m_nHeight = m_Parse.attrInt(_T("Height"), 300);
				_pWorkBook->m_nMsoCTPDockPosition = (MsoCTPDockPosition)m_Parse.attrInt(_T("DockPosition"), 4);
				_pWorkBook->m_nMsoCTPDockPositionRestrict = (MsoCTPDockPositionRestrict)m_Parse.attrInt(_T("DockPositionRestrict"), 3);
				vector<CTangramXmlParse*> vec;
				int nCount = m_Parse.GetCount();
				for (int i = nCount - 1; i >= 0; i--)
				{
					CTangramXmlParse* pChild = m_Parse.GetChild(i);
					CString strName = pChild->name();
					if (strName.CompareNoCase(_T("userforms")) && strName.CompareNoCase(_T("default")))
					{
						CString strType = pChild->attr(_T("type"), _T(""));
						XlSheetType nType = xlWorksheet;
						if (strType.CompareNoCase(_T("Chart")) == 0)
						{
							nType = xlChart;
						}
						CTangramXmlParse* pChild2 = pChild->GetChild(_T("default"));
						if (pChild2)
						{
							pChild2 = pChild2->GetChild(_T("sheet"));
							CComPtr<IDispatch> pDisp;
							if (nType == xlWorksheet)
							{
								//pDisp = pActive;
							}

							if (pDisp == nullptr)
								pSheets->Add(vtMissing, vtMissing, vtMissing, CComVariant(nType), 0, &pDisp);

							CComQIPtr<Excel::_Worksheet> pSh(pDisp);
							CString _strName = pChild2->attr(_T("caption"), _T(""));
							if (_strName == _T(""))
								_strName = pChild2->name();

							if (pSh)
							{
								pSh->put_Name(_strName.AllocSysString());
								CComPtr<Excel::CustomProperties> _pProps;
								pSh->get_CustomProperties(&_pProps);
								ICustomProperties* pProps = (ICustomProperties*)_pProps.p;
								CComPtr<Excel::CustomProperty> pProp;
								pProps->Add(CComBSTR(L"Tangram"), CComVariant(pChild->xml()), &pProp);
								vec.push_back(pChild);

								auto it = pExcelObj->m_pWorkBook->find(pSh.p);
								if (it == pExcelObj->m_pWorkBook->end())
								{
									CExcelWorkSheet* pExcelWorkSheet = new CExcelWorkSheet();
									pExcelWorkSheet->m_strSheetName = _strName;
									pExcelWorkSheet->m_pSheet = pSh.p;
									pExcelWorkSheet->m_pSheet->AddRef();
									(*_pWorkBook)[pExcelWorkSheet->m_pSheet] = pExcelWorkSheet;

									int nCount = pChild->GetCount();
									for (int i = 0; i < nCount; i++)
									{
										CTangramXmlParse* _pChild = pChild->GetChild(i);
										CString strName = _pChild->name();
										(*pExcelWorkSheet)[strName] = _pChild->xml();
									}
								}
							}
							else
							{
								CComQIPtr<Excel::_Chart> pChart(pDisp);
								if (pChart)
								{
									pChart->put_Name(_strName.AllocSysString());
								}
							}
						}
					}
					else if (strName.CompareNoCase(_T("userforms")) == 0)
					{
						int nCount = pChild->GetCount();
						for (int i = 0; i < nCount; i++)
						{
							CTangramXmlParse* _pChild = pChild->GetChild(i);
							CString _strName = _pChild->name();
							_pWorkBook->m_mapUserFormScript[_strName] = _pChild->xml();
						}
					}
					else
					{
						if (pChild)
						{
							_pWorkBook->m_strDefaultSheetXml = pChild->xml();
						}
					}
				}

				while (vec.size())
				{
					CTangramXmlParse* pParse = *vec.begin();
					m_Parse.RemoveNode(pParse);
					vec.erase(vec.begin());
				}

				_pWorkBook->m_strDocXml = m_Parse.xml();
				VARIANT_BOOL bRet = false;
				AddDocXml(pWorkBook, CComBSTR(_pWorkBook->m_strDocXml), CComBSTR(L"workbook"), &bRet);
				::PostMessage(m_hHostWnd, WM_OPENDOCUMENT, (WPARAM)pExcelObj, 0);
			}
			else
			{
				_pWorkBook->InitWorkBook();
			}
			//theApp.InitJava(0);
		}

		CExcelObject::CExcelObject(void)
		{
			m_hExcelEdit = NULL;
			m_hExcelEdit2 = NULL;
			m_pWorkBook = nullptr;
		}

		CExcelObject::~CExcelObject(void)
		{
		}

		void CExcelObject::SheetNameChanged()
		{
			CComPtr<IDispatch> pDisp;
			m_pWorkBook->m_pWorkBook->get_ActiveSheet(&pDisp);
			CString strName = theApp.GetPropertyFromObject(pDisp, _T("Name"));
			if (strName != m_strActiveSheetName)
			{
				ATLTRACE(_T("namechanged: oleName: %s, newName:%s\n"), m_strActiveSheetName, strName);
			}
		}

		void CExcelObject::OnOpenDoc()
		{
			if (m_pWorkBook->m_strDocXml == _T(""))
				return ;
			CExcelCloudAddin* pAddin = (CExcelCloudAddin*)theApp.m_pHostCore;
			pAddin->CreateWndPage((long)m_hClient, &m_pWorkBook->m_pPage);
			if (m_pWorkBook->m_pPage)
			{
				m_pWorkBook->m_pPage->put_External(m_pWorkBook->m_pWorkBook);
				IWndFrame* pFrame = nullptr;
				m_pWorkBook->m_pPage->CreateFrame(CComVariant(0), CComVariant((long)m_hChildClient), CComBSTR(L"Document"), &pFrame);
				m_pWorkBook->m_pFrame = (CWndFrame*)pFrame;
				if (m_pWorkBook->m_pFrame)
				{
					//if (m_bOldVer) 
					//	pExcelWorkBookWnd->m_pWorkBook->m_pFrame->ModifyHost((LONGLONG)m_hClientWnd);
					IDispatch* pDisp = nullptr;
					m_pWorkBook->m_pWorkBook->get_ActiveSheet(&pDisp);
					CComQIPtr<Excel::_Worksheet> pSheet(pDisp);

					CString strName = _T("");
					CString strXml = m_pWorkBook->m_strDefaultSheetXml;
					if (pSheet)
					{
						CComPtr<Excel::CustomProperties> pCustomProperties;
						pSheet->get_CustomProperties(&pCustomProperties);
						Excel::ICustomProperties* pProps = (Excel::ICustomProperties*)pCustomProperties.p;
						long nCount = 0;
						CComBSTR bstrName(L"");
						pProps->get_Count(&nCount);
						for (int i = 1; i <= nCount; i++)
						{
							CComPtr<Excel::CustomProperty> pProp;
							pProps->get_Item(CComVariant(i), &pProp);
							Excel::ICustomProperty* _pProp = (ICustomProperty*)pProp.p;
							_pProp->get_Name(&bstrName);
							strName = OLE2T(bstrName);
							if (strName == _T("Tangram"))
							{
								CComVariant var;
								_pProp->get_Value(&var);
								strXml = OLE2T(var.bstrVal);
								::VariantClear(&var);
								break;
							}
						}
					}

					if (strXml == _T(""))
						return ;
					CTangramXmlParse m_Parse;
					bool b = m_Parse.LoadXml(strXml);
					if (b)
					{
						CTangramXmlParse* pParse = m_Parse.GetChild(_T("default"));
						if (pParse)
						{
							strName = m_Parse.name();
							CString strKey = strName;
							strName += _T("_default");
							CTangramXmlParse* pParse2 = pParse->GetChild(_T("sheet"));
							IWndNode* pNode = nullptr;
							m_pWorkBook->m_pFrame->Extend(strName.AllocSysString(), CComBSTR(pParse2->xml()), &pNode);
							CWndNode* _pNode = (CWndNode*)pNode;
							if (_pNode->m_pOfficeObj == nullptr&&pDisp)
							{
								_pNode->m_pOfficeObj = pDisp;
								_pNode->m_pOfficeObj->AddRef();
							}
						}
						pParse = pParse->GetChild(_T("taskpane"));
						if (pParse)
						{
							strName += _T("_taskpane");

							Office::_CustomTaskPane* _pCustomTaskPane = nullptr;
							CString strSheetName = _T("");
							auto it = pAddin->m_mapTaskPaneMap.find((long)m_hForm);
							if (it == pAddin->m_mapTaskPaneMap.end())
							{
								VARIANT vWindow;
								vWindow.vt = VT_DISPATCH;
								vWindow.pdispVal = nullptr;
								HRESULT hr = pAddin->m_pCTPFactory->CreateCTP(L"Tangram.Ctrl.1", m_pWorkBook->m_strTaskPaneTitle.AllocSysString(), vWindow, &_pCustomTaskPane);
								_pCustomTaskPane->AddRef();
								pAddin->m_mapTaskPaneMap[(long)m_hForm] = _pCustomTaskPane;
								CComPtr<ITangramCtrl> pCtrlDisp;
								_pCustomTaskPane->get_ContentControl((IDispatch**)&pCtrlDisp);
								if (pCtrlDisp)
								{
									long hWnd = 0;
									pCtrlDisp->get_HWND(&hWnd);
									m_hTaskPane = (HWND)hWnd;
									if (m_pWorkBook->m_pTaskPanePage == nullptr)
									{
										HWND hPWnd = ::GetParent((HWND)hWnd);
										CWindow m_Wnd;
										m_Wnd.Attach(hPWnd);
										m_Wnd.ModifyStyle(0, WS_CLIPSIBLINGS | WS_CLIPCHILDREN);
										m_Wnd.Detach();
										HRESULT hr = theApp.m_pHostCore->CreateWndPage((long)hPWnd, &m_pWorkBook->m_pTaskPanePage);
										if (m_pWorkBook->m_pTaskPanePage)
										{
											IWndFrame* pTaskPaneFrame = nullptr;
											m_pWorkBook->m_pTaskPanePage->CreateFrame(CComVariant(0), CComVariant((long)hWnd), CComBSTR(L"TaskPane"), &pTaskPaneFrame);
											m_pWorkBook->m_pTaskPaneFrame = (CWndFrame*)pTaskPaneFrame;
											if (m_pWorkBook->m_pTaskPaneFrame)
											{
												IWndNode* pNode = nullptr;
												m_pWorkBook->m_pTaskPaneFrame->Extend(CComBSTR(strName), pParse->xml().AllocSysString(), &pNode);
												CWndNode* _pNode = (CWndNode*)pNode;
												if (_pNode->m_pOfficeObj == nullptr&&pDisp)
												{
													_pNode->m_pOfficeObj = pDisp;
													_pNode->m_pOfficeObj->AddRef();
												}
											}
										}
										_pCustomTaskPane->put_DockPosition(m_pWorkBook->m_nMsoCTPDockPosition);
										_pCustomTaskPane->put_DockPositionRestrict(m_pWorkBook->m_nMsoCTPDockPositionRestrict);
										_pCustomTaskPane->put_Width(m_pWorkBook->m_nWidth);
										_pCustomTaskPane->put_Height(m_pWorkBook->m_nHeight);
										_pCustomTaskPane->put_Visible(true);
									}
									else
										m_pWorkBook->m_pTaskPaneFrame->ModifyHost(hWnd);
								}
							}
							else
							{
								_pCustomTaskPane = it->second;
								_pCustomTaskPane->put_Visible(true);
								if (m_pWorkBook->m_pTaskPaneFrame)
								{
									IWndNode* pNode = nullptr;
									m_pWorkBook->m_pTaskPaneFrame->Extend(CComBSTR(strName), pParse->xml().AllocSysString(), &pNode);
								}
							}
						}
					}
				}
			}
		}

		void CExcelObject::ProcessMouseDownMsg(HWND hWnd)
		{
			HWND hActiveWnd = ::GetActiveWindow();
			if (hActiveWnd == m_hForm)
			{
				if (theApp.m_pWndNode && (::IsChild(theApp.m_pWndNode->m_pHostWnd->m_hWnd, hWnd) || hWnd == theApp.m_pWndNode->m_pHostWnd->m_hWnd))
				{
					if (theApp.m_pWndNode->m_nViewType == ViewType::Splitter || theApp.m_pWndNode->m_nViewType == ViewType::TreeView)
						return;
					if (::IsWindowEnabled(m_hExcelEdit))
						::EnableWindow(m_hExcelEdit, false);
					if (::IsWindowVisible(m_hExcelEdit2) || ::IsWindowEnabled(m_hExcelEdit2))
					{
						::PostMessage(m_hChildClient, WM_KEYDOWN, VK_TAB, 0);
						::PostMessage(m_hChildClient, WM_KEYDOWN, VK_LEFT, 0);
						::EnableWindow(m_hExcelEdit2, false);
						::PostMessage(theApp.m_pWndNode->m_pHostWnd->m_hWnd, WM_SETWNDFOCUSE, (WPARAM)hWnd, (LPARAM)m_hChildClient);
					}
				}
				else
				{
					::EnableWindow(m_hExcelEdit, true);
					::EnableWindow(m_hExcelEdit2, true);
					if (::IsWindowVisible(m_hChildClient) == false && theApp.m_pWndNode)
					{
						::PostMessage(theApp.m_pWndNode->m_pHostWnd->m_hWnd, WM_SETWNDFOCUSE, (WPARAM)hWnd, (LPARAM)m_hChildClient);
					}
				}
			}
			else
			{
				::EnableWindow(m_hExcelEdit, false);
				if (::IsWindowVisible(m_hExcelEdit2))
				{
					::PostMessage(m_hChildClient, WM_KEYDOWN, VK_TAB, 0);
					::PostMessage(m_hChildClient, WM_KEYDOWN, VK_LEFT, 0);
					::EnableWindow(m_hExcelEdit2, false);
				}
			}
		}

		void CExcelObject::OnObjDestory()
		{
			if (m_pOfficeObj != nullptr)
			{
				CExcelCloudAddin* pAddin = ((CExcelCloudAddin*)theApp.m_pHostCore);

				if (pAddin->m_pActiveExcelObject == this)
				{
					pAddin->m_pActiveExcelObject = nullptr;
					theApp.m_pWndNode = nullptr;
				}

				auto it2 = m_pWorkBook->m_mapExcelWorkBookWnd.find(m_hChildClient);
				if (it2 != m_pWorkBook->m_mapExcelWorkBookWnd.end())
				{
					m_pWorkBook->m_mapExcelWorkBookWnd.erase(it2);
				}

				size_t nCount = m_pWorkBook->m_mapExcelWorkBookWnd.size();
				if (nCount > 0 && m_pWorkBook->m_pFrame)
				{
					it2 = m_pWorkBook->m_mapExcelWorkBookWnd.begin();
					m_pWorkBook->m_pFrame->ModifyHost((long)::GetParent(it2->second->m_hChildClient));
					if (m_hTaskPane)
					{
						m_hTaskPane = NULL;
						m_pWorkBook->m_pTaskPaneFrame->ModifyHost((long)it2->second->m_hTaskPaneChildWnd);
						::DestroyWindow(m_hTaskPaneWnd);
					}
				}

				auto it = pAddin->m_mapTaskPaneMap.find((long)m_hForm);
				if (it != pAddin->m_mapTaskPaneMap.end())
				{
					it->second->Delete();
					pAddin->m_mapTaskPaneMap.erase(it);
				}
				if (nCount == 0)
				{
					HRESULT hr = S_OK;
					for (auto it : *m_pWorkBook)
					{
						delete it.second;
					}
					m_pWorkBook->erase(m_pWorkBook->begin(), m_pWorkBook->end());
					hr = m_pWorkBook->DispEventUnadvise(m_pWorkBook->m_pWorkBook);
					auto it2 = pAddin->find(m_pWorkBook->m_pWorkBook);
					if (it2 != pAddin->end())
						pAddin->erase(it2);
					ULONG dw = m_pWorkBook->m_pWorkBook->Release();
					while (dw>1)
						dw = m_pWorkBook->m_pWorkBook->Release();

					delete m_pWorkBook;
				}
				auto it3 = pAddin->m_mapOfficeObjects2.find(m_hForm);
				if (it3 != pAddin->m_mapOfficeObjects2.end())
					pAddin->m_mapOfficeObjects2.erase(it3);
			}
		}
	}
}
