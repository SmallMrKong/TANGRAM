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
#include "../../WndPage.h"
#include "../../WndNode.h"
#include "../../WndFrame.h"
#include "../../CloudAddinApp.h"
#include "../../TangramHtmlTreeWnd.h"
#include "ExcelPlusWnd.h"
#include "ExcelAddin.h"

namespace OfficeCloudPlus
{
	namespace ExcelPlus
	{
		CExcelWorkBook::CExcelWorkBook(void)
		{
#ifdef _DEBUG
			theApp.m_nOfficeDocs++;
#endif			
			m_pPage							= nullptr;
			m_pFrame						= nullptr;
			m_pWorkBook						= nullptr;
			m_pSheetNode					= nullptr;
			m_pTaskPanePage					= nullptr;
			m_pTaskPaneFrame				= nullptr;
			m_strDocXml						= _T("");
			m_strTaskPaneXml				= _T("");
			m_strDefaultSheetXml			= _T("");

		}

		CExcelWorkBook::~CExcelWorkBook(void)
		{
#ifdef _DEBUG
			theApp.m_nOfficeDocs--;
#endif			
		}

		void CExcelWorkBook::InitWorkBook()
		{
			CExcelCloudAddin* pAddin = (CExcelCloudAddin*)theApp.m_pHostCore;

			auto t = create_task([pAddin, this]()
			{
				CoInitializeEx(NULL, COINIT_MULTITHREADED);
				if (m_strDocXml==_T(""))
				{
					BSTR bstrXml(L"");
					pAddin->GetDocXmlByKey(m_pWorkBook, CComBSTR(L"workbook"), &bstrXml);
					m_strDocXml = OLE2T(bstrXml);
					CTangramXmlParse m_Parse;
					m_Parse.LoadXml(m_strDocXml);
					m_strTaskPaneTitle = m_Parse.attr(_T("Title"), _T("TaskPane"));
					m_nWidth = m_Parse.attrInt(_T("Width"), 200);
					m_nHeight = m_Parse.attrInt(_T("Height"), 300);
					m_nMsoCTPDockPosition = (MsoCTPDockPosition)m_Parse.attrInt(_T("DockPosition"), 4);
					m_nMsoCTPDockPositionRestrict = (MsoCTPDockPositionRestrict)m_Parse.attrInt(_T("DockPositionRestrict"), 3);

					CTangramXmlParse* pChild = m_Parse.GetChild(_T("default"));
					if (pChild == nullptr)
						return;
					m_strDefaultSheetXml = pChild->xml();
				}
				else
				{
					VARIANT_BOOL bRet = false;
					pAddin->AddDocXml(m_pWorkBook, CComBSTR(m_strDocXml), CComBSTR(L"workbook"), &bRet);
				}
				CoUninitialize();
			}).then([this]()
			{
				CString strKey = _T("<window>");
				CString strData = _T("");
				CString strVal = _T("");

				CString _strXml = theApp.GetXmlData(_T("userforms"), m_strDocXml);
				int nPos = _strXml.Find(strKey);
				while (nPos != -1)
				{
					strData = _strXml.Left(nPos);
					nPos = strData.ReverseFind('<');
					strData = strData.Mid(nPos + 1);
					nPos = strData.ReverseFind('>');
					strData = strData.Left(nPos);
					strData.Trim();
					strVal = theApp.GetXmlData(strData, _strXml);
					strKey = _T("</window>");
					nPos = _strXml.Find(strKey);
					_strXml = _strXml.Mid(nPos + 9);
					nPos = _strXml.Find(_T("<window>"));
					m_mapUserFormScript[strData] = strVal;
				}

				::PostMessage(theApp.m_pHostCore->m_hHostWnd, WM_OPENDOCUMENT, (WPARAM)m_mapExcelWorkBookWnd.begin()->second, 0);
			});
		}

		void CExcelWorkBook::ModifySheetForTangram(IDispatch* Sh, CString strSheetXml, CString strTaskPaneXml)
		{
			if (Sh == nullptr)
			{
				return;
			}

			CComBSTR szMember(L"CustomProperties");
			DISPID dispid = -1;
			HRESULT hr = Sh->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_USER_DEFAULT, &dispid);
			if (hr == S_OK)
			{
				DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
				VARIANT result = { 0 };
				EXCEPINFO excepInfo;
				memset(&excepInfo, 0, sizeof excepInfo);
				UINT nArgErr = (UINT)-1;
				HRESULT hr = Sh->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &result, &excepInfo, &nArgErr);
				if (S_OK == hr && VT_DISPATCH == result.vt)
				{
					CString strXml = _T("");
					if(strTaskPaneXml!=_T("")&& strSheetXml!=_T(""))
						strXml.Format(_T("<sheet><default><sheet><window>%s</window></sheet><taskpane><window>%s</window></taskpane></default></sheet>"), strSheetXml, strTaskPaneXml);
					else
					{
						if(strTaskPaneXml == _T("") && strSheetXml != _T(""))
							strXml.Format(_T("<sheet><default><sheet><window>%s</window></sheet></default></sheet>"), strSheetXml);
						else
						{
							strXml.Format(_T("<sheet><default><sheet><window><node name=\"Start\" id=\"hostview\" /></window></sheet><taskpane><window>%s</window></taskpane></default></sheet>"), strTaskPaneXml);
						}
					}
					CComQIPtr<Excel::CustomProperties> pCustomProperties(result.pdispVal);
					Excel::ICustomProperties* pProps = (Excel::ICustomProperties*)pCustomProperties.p;
					if (pCustomProperties)
					{
						CComPtr<Excel::CustomProperty> pProp;
						long nCount = 0;
						CComBSTR bstrName(L"");
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
								_pProp->put_Value(CComVariant(strXml));
								::VariantClear(&result);
								return;
							}
						}
					}
					CComPtr<Excel::CustomProperty> pProp;
					pProps->Add(CComBSTR(L"Tangram"), CComVariant(strXml), &pProp);
				}
				::VariantClear(&result);
			}
		}

		void CExcelWorkBook::OnNewSheet(IDispatch* Sh)
		{
			CComQIPtr<_Worksheet> pSheet(Sh);
			if(pSheet)
			{
				auto it = find(pSheet.p);
				if (it == end())
				{
					CComBSTR bstrName(L"");
					pSheet->get_Name(&bstrName);
					CExcelWorkSheet* pExcelWorkSheet = new CExcelWorkSheet();
					pExcelWorkSheet->m_strSheetName = OLE2T(bstrName);
					pExcelWorkSheet->m_pSheet = pSheet.p;
					pExcelWorkSheet->m_pSheet->AddRef();
					(*this)[pExcelWorkSheet->m_pSheet] = pExcelWorkSheet;

					//CComPtr<Excel::CustomProperties> _pProps;
					//pSheet->get_CustomProperties(&_pProps);
					//ICustomProperties* pProps = (ICustomProperties*)_pProps.p;
					//CComPtr<Excel::CustomProperty> pProp;
					//pProps->Add(CComBSTR(L"Tangram"), CComVariant(m_strDefaultSheetXml), &pProp);
				}
			}
		}

		void CExcelWorkBook::OnSheetBeforeDelete(IDispatch* Sh)
		{
			CComQIPtr<_Worksheet> pSheet(Sh);
			if(pSheet)
			{
				auto it = find(pSheet.p);
				if(it!=end())
				{
					delete it->second;
					erase(it);
				}
			}
		}

		void CExcelWorkBook::OnBeforeSave(VARIANT_BOOL SaveAsUI, VARIANT_BOOL* Cancel)
		{
			CExcelCloudAddin* pAddin = (CExcelCloudAddin*)theApp.m_pHostCore;
			if (m_pFrame)
			{
				m_pFrame->UpdateWndNode();
			}
			if (m_pTaskPaneFrame)
			{
				m_pTaskPaneFrame->UpdateWndNode();
			}
			HWND hPWnd = m_pFrame->GetTopLevelParent().m_hWnd;
			auto it1 = pAddin->m_mapTaskPaneMap.find((long)hPWnd);
			if (it1 != pAddin->m_mapTaskPaneMap.end())
			{
				BSTR bstrXml(L"");
				pAddin->GetDocXmlByKey(m_pWorkBook, CComBSTR(L"workbook"), &bstrXml);
				CString strXml = OLE2T(bstrXml);
				CTangramXmlParse m_Parse;
				bool b = m_Parse.LoadXml(strXml);
				if (b)
				{
					it1->second->get_Title(&bstrXml);
					strXml = OLE2T(bstrXml);
					m_Parse.put_attr(_T(""), strXml);
					int nValue = 0;
					it1->second->get_Width(&nValue);
					m_Parse.put_attr(_T("Width"), nValue);
					it1->second->get_Height(&nValue);
					m_Parse.put_attr(_T("Height"), nValue);
					Office::MsoCTPDockPosition m_Pos;
					it1->second->get_DockPosition(&m_Pos);
					m_Parse.put_attr(_T("DockPosition"), m_Pos);
					Office::MsoCTPDockPositionRestrict m_MsoCTPDockPositionRestrict;
					it1->second->get_DockPositionRestrict(&m_MsoCTPDockPositionRestrict);
					m_Parse.put_attr(_T("DockPositionRestrict"), m_MsoCTPDockPositionRestrict);
					VARIANT_BOOL bRet = false;
					pAddin->AddDocXml(m_pWorkBook, CComBSTR(m_Parse.xml()), CComBSTR(L"workbook"), &bRet);
				}
			}
		}

		void CExcelWorkBook::OnSheetActivate(IDispatch* Sh)
		{
			if (m_pFrame == nullptr)
				return;

			CExcelCloudAddin* pAddin = (CExcelCloudAddin*)theApp.m_pHostCore;
			CComQIPtr<Excel::_Worksheet> pSheet(Sh);
			CExcelObject* pExcelWorkBookWnd = pAddin->m_pActiveExcelObject;

			CString strName = _T("");
			CString strXml = m_strDefaultSheetXml;
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
				return;
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
					m_pFrame->Extend(strName.AllocSysString(), CComBSTR(pParse2->xml()), &pNode);
					CWndNode* _pNode = (CWndNode*)pNode;
					if (_pNode->m_pOfficeObj == nullptr)
					{
						_pNode->m_pOfficeObj = Sh;
						_pNode->m_pOfficeObj->AddRef();
					}
				}
				pParse = pParse->GetChild(_T("taskpane"));
				if (pParse)
				{
					strName += _T("_taskpane");

					Office::_CustomTaskPane* _pCustomTaskPane = nullptr;
					CString strSheetName = _T("");
					auto it = pAddin->m_mapTaskPaneMap.find((long)pExcelWorkBookWnd->m_hForm);
					if (it == pAddin->m_mapTaskPaneMap.end())
					{
						VARIANT vWindow;
						vWindow.vt = VT_DISPATCH;
						vWindow.pdispVal = nullptr;
						HRESULT hr = pAddin->m_pCTPFactory->CreateCTP(L"Tangram.Ctrl.1", m_strTaskPaneTitle.AllocSysString(), vWindow, &_pCustomTaskPane);
						_pCustomTaskPane->AddRef();
						pAddin->m_mapTaskPaneMap[(long)pAddin->m_pActiveExcelObject->m_hForm] = _pCustomTaskPane;
						_pCustomTaskPane->put_Visible(true);
						CComPtr<ITangramCtrl> pCtrlDisp;
						_pCustomTaskPane->get_ContentControl((IDispatch**)&pCtrlDisp);
						if (pCtrlDisp)
						{
							long hWnd = 0;
							pCtrlDisp->get_HWND(&hWnd);
							pExcelWorkBookWnd->m_hTaskPane = (HWND)hWnd;
							if (m_pTaskPanePage == nullptr)
							{
								HWND hPWnd = ::GetParent((HWND)hWnd);
								CWindow m_Wnd;
								m_Wnd.Attach(hPWnd);
								m_Wnd.ModifyStyle(0, WS_CLIPSIBLINGS | WS_CLIPCHILDREN);
								m_Wnd.Detach();
								HRESULT hr = theApp.m_pHostCore->CreateWndPage((long)hPWnd, &m_pTaskPanePage);
								if (m_pTaskPanePage)
								{
									IWndFrame* pTaskPaneFrame = nullptr;
									m_pTaskPanePage->CreateFrame(CComVariant(0), CComVariant((long)hWnd), CComBSTR(L"TaskPane"), &pTaskPaneFrame);
									m_pTaskPaneFrame = (CWndFrame*)pTaskPaneFrame;
									if (m_pTaskPaneFrame)
									{
										IWndNode* pNode = nullptr;
										m_pTaskPaneFrame->Extend(CComBSTR(strName), pParse->xml().AllocSysString(), &pNode);
										CWndNode* _pNode = (CWndNode*)pNode;
										if (_pNode->m_pOfficeObj == nullptr)
										{
											_pNode->m_pOfficeObj = Sh;
											_pNode->m_pOfficeObj->AddRef();
										}
									}
								}
							}
							else
								m_pTaskPaneFrame->ModifyHost(hWnd);
						}
					}
					else
					{
						_pCustomTaskPane = it->second;
						_pCustomTaskPane->put_Visible(true);
						if (m_pTaskPaneFrame)
						{
							IWndNode* pNode = nullptr;
							m_pTaskPaneFrame->Extend(CComBSTR(strName), pParse->xml().AllocSysString(), &pNode);
						}
					}
				}
			}
		}

		CExcelWorkSheet::CExcelWorkSheet(void)
		{
#ifdef _DEBUG
			theApp.m_nOfficeDocsSheet++;
#endif	
			m_pSheet			= nullptr;
			m_strKey			= _T("");
			m_strSheetName		= _T("");
		}

		CExcelWorkSheet::~CExcelWorkSheet(void)
		{
#ifdef _DEBUG
			theApp.m_nOfficeDocsSheet--;
#endif	
			m_pSheet			= nullptr;
			m_strKey			= _T("");
			m_strSheetName		= _T("");
		}
	}
}
