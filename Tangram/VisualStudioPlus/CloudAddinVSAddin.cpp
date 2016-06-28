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
*
********************************************************************************/

#include "../stdafx.h"
#include "../CloudAddinApp.h"
#include "../WndNode.h"
#include "../WndFrame.h"
#include "../TangramHtmlTreeWnd.h"
#include "CloudAddinVSAddin.h"
#include "dte.h"
#include "VsProject.h"

#ifndef _WIN64
namespace VSCloudPlus
{
	namespace VisualStudioPlus
	{
		CVSComponentWnd::CVSComponentWnd()
		{
			m_strName				= _T("");
			m_hHostWnd				= NULL;
			m_hChildHostWnd			= NULL;
			m_bCloudAddinComponent	= false;

			m_pNode					= nullptr;
			m_pFrame				= nullptr;
			m_pPage					= nullptr;
			m_pDesignerWnd			= nullptr;
		}

		LRESULT CVSComponentWnd::OnTangramMsg(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			HWND hWnd = (HWND)wParam;
			CString strName = (LPCTSTR)lParam;
			if (::IsWindow(hWnd))
			{
				CVSComponentWnd* pWnd = NULL;
				LRESULT lRes = ::SendMessage(hWnd, WM_TANGRAMDATA, 0, 0);
				if (lRes == 0)
				{
					pWnd = new CVSComponentWnd();
					pWnd->m_strName = strName;
					pWnd->SubclassWindow(hWnd);
					pWnd->m_pDesignerWnd = m_pDesignerWnd;
					pWnd->Init();
				}
				else
				{
					pWnd = (CVSComponentWnd*)lRes;
					pWnd->m_strName = strName;
				}
			}
			return DefWindowProc(uMsg, wParam, lParam);
		}

		LRESULT CVSComponentWnd::OnNameChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			CString strNewName = (LPCTSTR)wParam;
			return DefWindowProc(uMsg, wParam, lParam);
		}

		LRESULT CVSComponentWnd::OnGetDesignWnd(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			CVSCloudAddin* pAddin = (CVSCloudAddin*)theApp.m_pHostCore;
			switch (lParam)
			{
			case 0:
				{
					if (m_pPage)
					{
						if (m_pFrame)
						{
							m_pFrame->m_bDesignerState = true;
							auto it = m_pFrame->m_mapNode.begin();
							if (it != m_pFrame->m_mapNode.end())
							{
								CWndNode* pNode = m_pFrame->m_mapNode.begin()->second;
								CTangramXmlParse* pParse = pNode->m_pTangramParse;
								CString strXml = pParse->xml();
								if (theApp.m_pDocDOMTree)
								{
									theApp.m_pHostDesignUINode = pNode;
									theApp.m_pDocDOMTree->UpdateData(strXml);
									if (pAddin->m_pOutputWindowPane)
									{
										pAddin->m_pOutputWindowPane->Clear();
										pAddin->m_pOutputWindowPane->OutputString(strXml.AllocSysString());
									}
								}
							}
						}
						return (LRESULT)m_hWnd;
					}
				}
				break;
			case 1:
				{
					CComBSTR bstrName(L"");
					m_pDesignerWnd->m_pPrj->get_FullName(&bstrName);
					CString strPath = OLE2T(bstrName);
					int nPos = strPath.ReverseFind('\\');
					strPath = strPath.Left(nPos + 1);
					strPath += _T("tangramresource.xml");
					strPath.MakeLower();
					CString strXml = _T("");
					auto it = pAddin->m_mapPrjXml.find(strPath);
					if (it != pAddin->m_mapPrjXml.end())
					{
						//CTangramXmlParse* pForms = it->second->GetChild(_T("Forms"));
						//if (pForms == NULL)
						//	pForms = it->second->GetChild(_T("forms"));
						//if (pForms)
						//{
						//	CTangramXmlParse* pDesignerParse = pForms->GetChild(m_pDesignerWnd->m_pDocWnd->m_strDesignerName);
						//	if (pDesignerParse)
						//	{
						//		CWndNode* pNode = (CWndNode*)m_pNode;
						//		if (m_pNode)
						//		{
						//			CTangramXmlParse* pParse = pNode->m_pTangramParse;
						//			pForms->ReplaceNode(pDesignerParse, pParse, _T(""));
						//			if (pAddin->m_pOutputWindowPane)
						//			{
						//				pAddin->m_pOutputWindowPane->Clear();
						//				pAddin->m_pOutputWindowPane->OutputString(pForms->xml().AllocSysString());
						//			}
						//			//it->second->SaveFile(strPath);
						//		}
						//	}
						//}
					}
				}
				break;
			case 100:
				{
					int nRet = ::MessageBox(::GetActiveWindow(), _T("Do you want to design this Component ?"), _T("Tangram"), MB_YESNOCANCEL);
					if (nRet == IDYES)
					{
						CVSCloudAddin* pAddin = (CVSCloudAddin*)theApp.m_pHostCore;
						if (pAddin->UpdateProjectforCloudAddin(m_pDesignerWnd->m_pPrj))
						{
							CString strXml = _T("");
							m_pDesignerWnd->m_pActiveWnd = this;
							strXml.Format(_T("<%s><window><node name=\"Start\"/></window></%s>"), m_strName, m_strName);
							pAddin->OnTangramDesignMsg(strXml);
						}
					}
				}
				break;
			}
			return DefWindowProc(uMsg, wParam, lParam);
		}

		LRESULT CVSComponentWnd::OnMouseActivate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			if (theApp.m_pCloudAddinCLRProxy&&m_pDesignerWnd->m_hWnd!=m_hWnd)
			{
				CString strInfo = theApp.m_pCloudAddinCLRProxy->GetCtrlInfo((long)m_hWnd);
				if (strInfo != _T(""))
				{
					int nPos = strInfo.Find(_T("|"));
					CString strOldName = m_strName;
					m_strName = strInfo.Left(nPos);
					if (m_strName != strOldName)
					{
						//Need Modify
					}
					BOOL bDock = _wtoi(strInfo.Mid(nPos + 1));

					if (bDock&&m_bCloudAddinComponent==false)
					{
						::PostMessage(m_hWnd, WM_TANGRAMGETDESIGNWND, 0, 100);
						//int nRet = ::MessageBox(NULL, _T("Do you want to design this Component ?"), _T("Tangram"), MB_YESNOCANCEL);
						//if (nRet == IDYES)
						//{
						//	CVSCloudAddin* pAddin = (CVSCloudAddin*)theApp.m_pHostCore;
						//	if (pAddin->UpdateProjectforCloudAddin(m_pDesignerWnd->m_pPrj))
						//	{
						//		CString strXml = _T("");
						//		m_pDesignerWnd->m_pActiveWnd = this;
						//		strXml.Format(_T("<%s><window><node name=\"Start\"/></window></%s>"), m_strName,m_strName);
						//		pAddin->OnTangramDesignMsg(strXml);
						//	}
						//}
					}
				}
			}
			return 1;
		}

		void CVSComponentWnd::Init()
		{
			BOOL bIsCloudAddinPrj = false;
			CVSCloudAddin* pAddin = (CVSCloudAddin*)theApp.m_pHostCore;
			if (m_pDesignerWnd&&m_pDesignerWnd->m_pPrj)
			{
				bIsCloudAddinPrj = pAddin->IsCloudAddinProject(m_pDesignerWnd->m_pPrj);
			}
			if (bIsCloudAddinPrj&&theApp.m_pCloudAddinCLRProxy&&m_pDesignerWnd->m_hWnd != m_hWnd)
			{
				CString strInfo = theApp.m_pCloudAddinCLRProxy->GetCtrlInfo((long)m_hWnd);
				if (strInfo != _T(""))
				{
					int nPos = strInfo.Find(_T("|"));
					m_strName = strInfo.Left(nPos);
					BOOL bDock = _wtoi(strInfo.Mid(nPos + 1));

					if (bDock)
					{
						CComBSTR bstrName(L"");
						m_pDesignerWnd->m_pPrj->get_FullName(&bstrName);
						CString strPath = OLE2T(bstrName);
						int nPos = strPath.ReverseFind('\\');
						strPath = strPath.Left(nPos + 1);
						strPath += _T("tangramresource.xml");
						strPath.MakeLower();
						CString strXml = _T("");
						CTangramXmlParse* pParse = NULL;
						auto it = pAddin->m_mapPrjXml.find(strPath);
						if (it != pAddin->m_mapPrjXml.end())
						{
							pParse = it->second;
						}
						else
							return;
						if (pParse)
						{
							CTangramXmlParse* pForms = pParse->GetChild(_T("Forms"));
							if (pForms == nullptr)
								pForms = pParse->GetChild(_T("forms"));
							if (pForms)
							{
								CTangramXmlParse* pDesignerParse = pForms->GetChild(m_pDesignerWnd->m_strName);
								if (pDesignerParse)
								{
									CTangramXmlParse* pChild = pDesignerParse->GetChild(m_strName);
									if (pChild)
									{
										strXml = pChild->xml();
										m_bCloudAddinComponent = true;
										m_pDesignerWnd->m_pActiveWnd = this;
										if(strXml!=_T(""))
										{
											CString strNewName = _T("");
											CTangramXmlParse* pParse = nullptr;
											CString strInfo = theApp.m_pCloudAddinCLRProxy->GetCtrlInfo((long)m_pDesignerWnd->m_hWnd);
											if (strInfo != _T(""))
											{
												int nPos = strInfo.Find(_T("|"));
												CString strName = strInfo.Left(nPos);
												if (m_pDesignerWnd->m_strName != strName)
												{
													m_pDesignerWnd->m_strName = strName;
												}
											}
											IWndFrame* pFrame = nullptr;
											HWND hParent = ::GetParent(m_hWnd);
											pAddin->CreateWndPage((long)hParent, &m_pPage);
											m_pPage->CreateFrame(CComVariant(0), CComVariant((long)m_hWnd), CComBSTR(m_strName), &pFrame);
											if (pFrame)
											{
												m_pFrame = (CWndFrame*)pFrame;
												m_pFrame->m_bDesignerState = true;
												m_pFrame->Extend(CComBSTR(L""), CComBSTR(strXml), &m_pNode);
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}

		void CVSComponentWnd::OnFinalMessage(HWND hWnd)
		{
			CWindowImpl::OnFinalMessage(hWnd);
			delete this;
		}

		CVSDesignerWnd::CVSDesignerWnd()
		{
			m_strName				= _T("");
			m_strPath				= _T("");
			m_bDesignState			= false;
			m_pPage					= nullptr;
			m_pFrame				= nullptr;
			m_pDocument				= nullptr;
			m_pDocEvent				= nullptr;
			m_pActiveWnd			= nullptr;
			m_hMdiClient			= nullptr;
		}

		void CVSDesignerWnd::OnDocumentSaved(VxDTE::Document* Document)
		{
			CTangramXmlParse* m_pPrjDocNode=nullptr;
			if (m_pPrjDocNode == nullptr)
			{
				CComBSTR bstrName(L"");
				m_pPrj->get_FullName(&bstrName);
				m_strPath = OLE2T(bstrName);
				int nPos = m_strPath.ReverseFind('\\');
				m_strPath = m_strPath.Left(nPos + 1);
				m_strPath += _T("tangramresource.xml");
				m_strPath.MakeLower();
				CVSCloudAddin* pAddin = (CVSCloudAddin*)theApp.m_pHostCore;
				auto it = pAddin->m_mapPrjXml.find(m_strPath);
				if (it != pAddin->m_mapPrjXml.end())
				{
					m_pPrjDocNode = it->second;
				}
			}
			CString strXml = _T("");
			CString strData = _T("");
			CString strInfo = theApp.m_pCloudAddinCLRProxy->GetCtrlInfo((long)m_hWnd);
			if (strInfo != _T(""))
			{
				m_strName = strInfo.Left(strInfo.Find(_T("|")));
			}

			CString strName = m_strName;
			if (::IsWindow(m_hMdiClient) && m_pNode)
			{
				theApp.m_pHostCore->UpdateWndNode(m_pNode);
				CWndNode* pNode = (CWndNode*)m_pNode;
				if (pNode->m_pTangramParse)
				{
					strXml.Format(_T("<%s>%s</%s>"), _T("mdiclient"), pNode->m_pTangramParse->GetChild(_T("window"))->xml(), _T("mdiclient"));
					strData += strXml;
				}
			}
			int nSize = size();
			if (nSize)
			{
				for (auto it : *this)
				{
					CVSComponentWnd* pWnd = it.second;
					strInfo = theApp.m_pCloudAddinCLRProxy->GetCtrlInfo((long)pWnd->m_hWnd);
					if (strInfo != _T(""))
					{
						pWnd->m_strName = strInfo.Left(strInfo.Find(_T("|")));
					}
					if (pWnd->m_pNode)
					{
						theApp.m_pHostCore->UpdateWndNode(pWnd->m_pNode);
						CWndNode* pNode = (CWndNode*)pWnd->m_pNode;
						if (pNode->m_pTangramParse)
						{
							strXml.Format(_T("<%s>%s</%s>"), pWnd->m_strName, pNode->m_pTangramParse->GetChild(_T("window"))->xml(), pWnd->m_strName);
							strData += strXml;
						}
					}
				}
			}
			CString str = _T("");
			str.Format(_T("<%s>%s</%s>"), strName, strData, strName);
			CTangramXmlParse m_Parse;
			if (m_Parse.LoadXml(str))
			{
				CTangramXmlParse* pParse = m_pPrjDocNode->GetChild(_T("Forms"));
				CTangramXmlParse* pParse2 = pParse->GetChild(strName);
				if (pParse2 == nullptr)
				{
					pParse2 = pParse->AddNode(strName);
				}
				pParse->ReplaceNode(pParse2, &m_Parse, _T(""));
				m_pPrjDocNode->SaveFile(m_strPath);
			}
		}

		void CVSDesignerWnd::OnDocumentClosing(VxDTE::Document* Document)
		{
		}

		void CVSDesignerWnd::OnDocumentOpening(BSTR DocumentPath, VARIANT_BOOL ReadOnly)
		{
		}

		void CVSDesignerWnd::OnDocumentOpened(VxDTE::Document* Document)
		{
		}

		LRESULT CVSDesignerWnd::OnNameChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			CString strNewName = (LPCTSTR)wParam;
			return DefWindowProc(uMsg, wParam, lParam);
		}

		LRESULT CVSDesignerWnd::OnMouseActivate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			if (m_bDesignState == false)
				::PostMessage(m_hWnd, WM_TANGRAMGETDESIGNWND, 0, 1);
			return DefWindowProc(uMsg, wParam, lParam);
		}

		LRESULT CVSDesignerWnd::OnGetDesignWnd(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			if (lParam == 1)
			{
				if (m_bDesignState == false&&::IsWindow(m_hMdiClient))
				{
					if (::MessageBox(::GetActiveWindow(), _T("Do you want to design this MDI Form?"), _T("Tangram"), MB_YESNOCANCEL) == IDYES)
					{
						Init();
					}
				}
			}
			return DefWindowProc(uMsg, wParam, lParam);
		}

		void CVSDesignerWnd::Init()
		{
			BOOL bIsCloudAddinPrj = false;
			CVSCloudAddin* pAddin = (CVSCloudAddin*)theApp.m_pHostCore;
			if (m_pPrj)
			{
				bIsCloudAddinPrj = pAddin->IsCloudAddinProject(m_pPrj);
				if (bIsCloudAddinPrj == false)
				{
					if (pAddin->UpdateProjectforCloudAddin(m_pPrj))
					{
						//CString strXml = _T("");
						//m_pDesignerWnd->m_pActiveWnd = this;
						//strXml.Format(_T("<%s><window><node name=\"Start\"/></window></%s>"), m_strName, m_strName);
						//pAddin->OnTangramDesignMsg(strXml);
					}
					//return;
				}
			}
			m_bDesignState = true;
			CComBSTR bstrName(L"");
			m_pPrj->get_FullName(&bstrName);
			CString strPath = OLE2T(bstrName);
			int nPos = strPath.ReverseFind('\\');
			strPath = strPath.Left(nPos + 1);
			strPath += _T("tangramresource.xml");
			strPath.MakeLower();
			CString strXml = _T("");
			auto it = pAddin->m_mapPrjXml.find(strPath);
			if (it != pAddin->m_mapPrjXml.end())
			{
				CTangramXmlParse* pForms = it->second->GetChild(_T("Forms"));
				if (pForms == nullptr)
					pForms = it->second->GetChild(_T("forms"));
				if (pForms)
				{
					CTangramXmlParse* pDesignerParse = pForms->GetChild(m_strName);
					if (pDesignerParse)
					{
						CTangramXmlParse* pChild = pDesignerParse->GetChild(_T("mdiclient"));
						if (pChild)
							strXml = pChild->xml();
					}
					if (strXml == _T(""))
					{
						strXml = _T("");
						strXml.Format(_T("<%s><window><node name=\"Start\"/></window></%s>"), _T("mdiclient"), _T("mdiclient"));
					}
				}
			}

			pAddin->CreateWndPage((long)m_hWnd, &m_pPage);
			IWndFrame* pFrame = NULL;
			m_pPage->CreateFrame(CComVariant(0), CComVariant((long)m_hMdiClient), CComBSTR(m_strName), &pFrame);
			if (pFrame)
			{
				m_pFrame = (CWndFrame*)pFrame;
				m_pFrame->put_DesignerState(true);
				m_pFrame->Extend(CComBSTR(L""), CComBSTR(strXml), &m_pNode);
				CWndNode* pNode = m_pFrame->m_pWorkNode;
				CTangramXmlParse* pParse = pNode->m_pTangramParse;
				CString strXml = pParse->xml();
				if (theApp.m_pDocDOMTree)
				{
					theApp.m_pDocDOMTree->DeleteItem(theApp.m_pDocDOMTree->m_hFirstRoot);
					if (theApp.m_pDocDOMTree->m_pHostXmlParse)
					{
						delete theApp.m_pDocDOMTree->m_pHostXmlParse;
					}
					theApp.InitDesignerTreeCtrl(strXml);
				}
			}
		}

		LRESULT CVSDesignerWnd::OnTangramMsg(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
		{
			HWND hWnd = (HWND)wParam;
			switch (lParam)
			{
				case 1992:
				{
					return 0;
				}
				break;
			case 1965:
				{
					return 0;
				}
				break;
			case 1963:
				{
					for (auto it = begin(); it != end(); it++)
					{
						if (it->second->m_hWnd == hWnd)
						{
							erase(it);
							break;
						}
					}
				}
				break;
			default:
				{
					if (::IsWindow(hWnd))
					{
						LRESULT lRes = ::SendMessage(hWnd, WM_TANGRAMDATA, 0, 0);
						if (lRes == 0)
						{
							if (m_hWnd != hWnd)
							{
								CVSComponentWnd* pWnd = new CVSComponentWnd();
								pWnd->m_strName = (LPCTSTR)lParam;
								pWnd->SubclassWindow(hWnd);
								pWnd->m_pDesignerWnd = this;
								(*this)[pWnd->m_strName.MakeLower()] = pWnd;
								pWnd->Init();
							}
						}
						else
						{
							CVSComponentWnd* pWnd = (CVSComponentWnd*)lRes;
							pWnd->m_strName = (LPCTSTR)lParam;
						}
					}
				}
				break;
			}
			
			return DefWindowProc(uMsg, wParam, lParam);
		}

		void CVSDesignerWnd::OnFinalMessage(HWND hWnd)
		{
			CVSCloudAddin* pAddin = (CVSCloudAddin*)theApp.m_pHostCore;
			auto it = pAddin->find(m_pDocument);
			if (it != pAddin->end())
				pAddin->erase(it);
			HRESULT hr = S_FALSE;
			hr = DispEventUnadvise((_DocumentEvents*)m_pDocEvent);
			CWindowImpl::OnFinalMessage(hWnd);
			delete this;
		}

		CVSCloudAddin::CVSCloudAddin()
		{
			m_hPropertyForm						= NULL;
			m_pDTE								= nullptr;
			m_pWindows							= nullptr;
			m_pActiveDesignerWnd				= nullptr;
			m_pPropertyWnd						= nullptr;
			m_pOutputWindowPane					= nullptr;
			m_pOutputWindowPane					= nullptr;
			m_pPropertyWndVisibilityEvents		= nullptr;
			m_strprjKindVBProject				= CString(VsProject::prjKindVBProject);
			m_strprjKindCSharpProject			= CString(VsProject::prjKindCSharpProject);
		}

		CVSCloudAddin::~CVSCloudAddin()
		{
			//if (m_pDTE)
			//	m_pDTE->Release();
			if (m_pWindows)
				m_pWindows->Release();
		}

		LRESULT CVSCloudAddin::OnTangramDesignMsg(CString _strXml)
		{
			//Need Modify
			CString strXml = _strXml;
			if (m_pActiveDesignerWnd)
			{
				BOOL bDesignerNameChange = false;
				CString strNewName = _T("");
				CTangramXmlParse* pParse = nullptr;
				if (theApp.m_pCloudAddinCLRProxy)
				{
					CString strInfo = theApp.m_pCloudAddinCLRProxy->GetCtrlInfo((long)m_pActiveDesignerWnd->m_hWnd);
					if (strInfo != _T(""))
					{
						int nPos = strInfo.Find(_T("|"));
						CString strName = strInfo.Left(nPos);
						if (m_pActiveDesignerWnd->m_strName != strName)
						{
							bDesignerNameChange = true;
							strNewName = strName;
						}
					}
				}
				CVSComponentWnd* pWnd = m_pActiveDesignerWnd->m_pActiveWnd;
				if (pWnd)
				{
					pWnd->m_bCloudAddinComponent = true;
					IWndFrame* pFrame = nullptr;
					HWND hParent = ::GetParent(pWnd->m_hWnd);
					CreateWndPage((long)hParent, &pWnd->m_pPage);
					pWnd->m_pPage->CreateFrame(CComVariant(0), CComVariant((long)pWnd->m_hWnd), CComBSTR(pWnd->m_strName), &pFrame);
					if (pFrame)
					{
						pWnd->m_pFrame = (CWndFrame*)pFrame;
						IWndNode* pNode = nullptr;
						pWnd->m_pFrame->Extend(CComBSTR(L""), CComBSTR(strXml), &pWnd->m_pNode);
						if (pWnd->m_pNode)
						{
							pWnd->m_pFrame->m_bDesignerState = true;
							if (m_pActiveDesignerWnd->m_pPrj)
							{
								CString _xmlStr = _T("<window><node name=\"Start\"/></window>");
								CComBSTR bstrName(L"");
								m_pActiveDesignerWnd->m_pPrj->get_FullName(&bstrName);
								CString strPath = OLE2T(bstrName);
								int nPos = strPath.ReverseFind('\\');
								strPath = strPath.Left(nPos+1);
								strPath += _T("tangramresource.xml");
								strPath.MakeLower();
								auto it = m_mapPrjXml.find(strPath);
								if (it != m_mapPrjXml.end())
								{
									pParse = it->second;
									CTangramXmlParse* pFormsParse = pParse->GetChild(_T("Forms"));
									if (pFormsParse == nullptr)
										pFormsParse = pParse->GetChild(_T("forms"));
									CTangramXmlParse* pDesignerParse = pFormsParse->GetChild(m_pActiveDesignerWnd->m_strName);
									if (pDesignerParse == nullptr)
									{
										pDesignerParse = pFormsParse->AddNode(m_pActiveDesignerWnd->m_strName);
									}
									CTangramXmlParse* pChild = pDesignerParse->GetChild(pWnd->m_strName);
									if (pChild)
									{
										pDesignerParse->RemoveNode(pWnd->m_strName);
									}
									pChild = pDesignerParse->AddNode(pWnd->m_strName);
									CString strGUID = theApp.GetNewGUID();
									pChild->put_text(strGUID);
									CString strNewXml = pParse->xml();
									strNewXml.Replace(strGUID, _xmlStr);
									pParse->Reflash();
									if (pParse->LoadXml(strNewXml))
									{
										pFormsParse = pParse->GetChild(_T("Forms"));
										if (pFormsParse == nullptr)
											pFormsParse = pParse->GetChild(_T("forms"));
										pDesignerParse = pFormsParse->GetChild(m_pActiveDesignerWnd->m_strName);
										if (bDesignerNameChange)
										{
											CString strDesignerXml = pDesignerParse->xml();
											nPos = strDesignerXml.Find(_T(">"));
											strDesignerXml = strDesignerXml.Mid(nPos + 1);
											nPos = strDesignerXml.ReverseFind('<');
											strDesignerXml = strDesignerXml.Left(nPos);
											pFormsParse->RemoveNode(pDesignerParse);
											pChild = pFormsParse->AddNode(strNewName);
											strGUID = theApp.GetNewGUID();
											pChild->put_text(strGUID);
											strNewXml = pParse->xml();
											strNewXml.Replace(strGUID, strDesignerXml);
											pParse->Reflash();
											pParse->LoadXml(strNewXml);
											m_pActiveDesignerWnd->m_strName = strNewName;
										}
									}
								}
							}
						}
						return 1;
					}
				}
			}

			return 0;
		}

		void CVSCloudAddin::PutTextToOutputPane(BSTR bstrMsg)
		{
			if (m_pOutputWindowPane)
			{
				m_pOutputWindowPane->OutputString(bstrMsg);
			}
		}

		BOOL CVSCloudAddin::IsCloudAddinProject(VxDTE::Project* pPrj)
		{
			if (pPrj)
			{
				CComBSTR bstrKind(L"");
				pPrj->get_Kind(&bstrKind);
				CString strKind = OLE2T(bstrKind);
				if (strKind.CompareNoCase(m_strprjKindCSharpProject) == 0 || strKind.CompareNoCase(m_strprjKindVBProject) == 0)
				{
					CComPtr<IDispatch> pDisp;
					pPrj->get_Object(&pDisp);
					CComQIPtr<VsProject::VSProject> pVSProject(pDisp);
					CComPtr<VsProject::Reference> pReference;
					CComPtr<VsProject::References> pReferences;
					HRESULT hr = pVSProject->get_References(&pReferences);
					if (pReferences)
					{
						long nCount = 0;
						pReferences->get_Count(&nCount);
						if (nCount)
						{
							for (int i = nCount; i >= 1; i--)
							{
								CComPtr<VsProject::Reference> pReference;
								pReferences->Item(CComVariant(i), &pReference);
								if (pReference)
								{
									CComBSTR bstrPath;
									pReference->get_Path(&bstrPath);
									CString strPath = OLE2T(bstrPath);
									bstrPath.Detach();
									if (strPath.CompareNoCase(theApp.m_strTangramCLRPath) == 0)
									{
										pPrj->get_FullName(&bstrKind);
										CString strPath = OLE2T(bstrKind);
										int nPos = strPath.ReverseFind('\\');
										strPath = strPath.Left(nPos + 1) + _T("tangramresource.xml");
										strPath.MakeLower();
										auto it = m_mapPrjXml.find(strPath);
										if (it == m_mapPrjXml.end())
										{
											if (::PathFileExists(strPath))
											{
												CTangramXmlParse* pParse = new CTangramXmlParse();
												if (pParse->LoadFile(strPath))
												{
													m_mapPrjXml[strPath] = pParse;
												}
												else
												{
													delete pParse;
													return false;
												}
											}
										}
										return true;
									}
								}
							}
						}
					}
				}
			}
			return false;
		}

		BOOL CVSCloudAddin::UpdateProjectforCloudAddin(VxDTE::Project* pPrj)
		{
			if (pPrj)
			{
				CComBSTR bstrKind(L"");
				pPrj->get_FullName(&bstrKind);
				CString strPath = OLE2T(bstrKind);
				int nPos = strPath.ReverseFind('\\');
				strPath = strPath.Left(nPos + 1) + _T("tangramresource.xml");
				strPath.MakeLower();
				auto it = m_mapPrjXml.find(strPath);
				if (it != m_mapPrjXml.end())
					return true;
				pPrj->get_Kind(&bstrKind);
				CString strKind = OLE2T(bstrKind);
				if (strKind.CompareNoCase(m_strprjKindCSharpProject) == 0 || strKind.CompareNoCase(m_strprjKindVBProject) == 0)
				{
					CComPtr<IDispatch> pDisp;
					pPrj->get_Object(&pDisp);
					CComQIPtr<VsProject::VSProject> pVSProject(pDisp);
					CComPtr<VsProject::Reference> pReference;
					CComPtr<VsProject::References> pReferences;
					HRESULT hr = pVSProject->get_References(&pReferences);
					pReferences->Add(theApp.m_strTangramCLRPath.AllocSysString(), &pReference);
					CTangramXmlParse* pParse = new CTangramXmlParse();
					if (!pParse->LoadFile(strPath))
					{
						CString strXml = _T("<Tangram><Forms></Forms></Tangram>");
						pParse->LoadXml(strXml);
						pParse->SaveFile(strPath);
					}
					m_mapPrjXml[strPath] = pParse;
					nPos = strPath.ReverseFind('\\');
					CString strName = strPath.Mid(nPos + 1);
					CComPtr<ProjectItem> pItem;
					CComPtr<ProjectItems> pItems;
					pPrj->get_ProjectItems(&pItems);
					pItems->Item(CComVariant(strName), &pItem);
					if (!pItem)
					{
						pItems->AddFromFile(CComBSTR(strPath), &pItem);
						if (hr == S_OK)
						{
							CComBSTR bstrName(L"");
							pItem->get_Name(&bstrName);
							strName = OLE2T(bstrName);
							CComPtr<VxDTE::Properties> pProperties;
							pItem->get_Properties(&pProperties);
							CComPtr<VxDTE::Property> pProperty;
							pProperties->Item(CComVariant(L"ItemType"), &pProperty);
							if (pProperty)
							{
								pProperty->put_Value(CComVariant(L"EmbeddedResource"));
							}
						}
					}
					return true;
				}
			}
			return false;
		}

		void CVSCloudAddin::AttachDoc(VxDTE::Document* pDoc)
		{
			if (pDoc)
			{
				auto it = find(pDoc);
				if (it == end())
				{
					CComPtr<VxDTE::Window> pWnd;
					pDoc->get_ActiveWindow(&pWnd);
					if (pWnd == nullptr)
						return;
					CComBSTR bstrKind(L"");
					pDoc->get_Kind(&bstrKind);
					CString strDocEditor = OLE2T(bstrKind);
					if (strDocEditor.CompareNoCase(CString(VxDTE::vsDocumentKindText)) == 0)
					{
						long h = 0;
						pWnd->get_HWnd(&h);
						if (h)
						{
							HWND hWnd = (HWND)h;
							m_pActiveDesignerWnd = nullptr;
							hWnd = ::GetWindow(hWnd, GW_CHILD);
							hWnd = ::GetWindow(hWnd, GW_CHILD);
							hWnd = ::FindWindowEx(hWnd, NULL, NULL, _T("OverlayControl"));
							if (hWnd)
							{
								hWnd = ::GetWindow(hWnd, GW_CHILD);
								hWnd = ::GetWindow(hWnd, GW_HWNDNEXT);
								::GetWindowText(hWnd, theApp.m_szBuffer, _MAX_FNAME);
								CString strTxt = CString(theApp.m_szBuffer);
								if (strTxt.CompareNoCase(_T("ToolStripAdornerWindow")) == 0)
								{
									hWnd = ::GetWindow(hWnd, GW_HWNDNEXT);
								}
								if (hWnd == NULL)
									return;

								m_pActiveDesignerWnd = new CVSDesignerWnd();
								m_pActiveDesignerWnd->SubclassWindow(hWnd);
								(*this)[pDoc] = m_pActiveDesignerWnd;
								m_pActiveDesignerWnd->m_pDocument = pDoc;
								m_pActiveDesignerWnd->m_pDocument->AddRef();

								CComPtr<Events> pEvents;
								HRESULT hr = m_pDTE->get_Events(&pEvents);
								hr = pEvents->get_DocumentEvents((VxDTE::Document*)m_pActiveDesignerWnd->m_pDocument, &m_pActiveDesignerWnd->m_pDocEvent);
								hr = m_pActiveDesignerWnd->DispEventAdvise((_DocumentEvents*)m_pActiveDesignerWnd->m_pDocEvent);
								pWnd->get_Project(&m_pActiveDesignerWnd->m_pPrj);
								if (theApp.m_pCloudAddinCLRProxy)
								{
									m_pActiveDesignerWnd->m_hMdiClient = theApp.m_pCloudAddinCLRProxy->OnInitialDesigner((long)hWnd, m_pActiveDesignerWnd->m_strName);
									if (m_pActiveDesignerWnd->m_hMdiClient)
									{
										m_pActiveDesignerWnd->Init();
									}
								}
							}
						}
					}
					if (m_hPropertyForm == NULL)
					{
						VARIANT_BOOL bVisible = false;
						m_pPropertyWnd->get_Visible(&bVisible);
						if (bVisible)
						{
							OnWindowShowing(nullptr);
						}
					}
				}
				else
				{
					m_pActiveDesignerWnd = it->second;
				}
			}
		}

		void CVSCloudAddin::InitPropertyGridCtrl(HWND hwnd)
		{
			if (m_hPropertyForm == NULL)
			{
				CComBSTR bstrCap(L"");
				m_pPropertyWnd->get_Caption(&bstrCap);
				m_hPropertyForm = ::FindWindowEx(hwnd, NULL, _T("GenericPane"), OLE2T(bstrCap));
				m_hPropertyForm = ::GetWindow(m_hPropertyForm, GW_CHILD);
				m_hPropertyForm = ::FindWindowEx(m_hPropertyForm, NULL, NULL, _T("PropertyGrid"));
				theApp.m_pCloudAddinCLRProxy->AttachPropertyGridCtrl(m_hPropertyForm);
			}
		}

		void CVSCloudAddin::OnInitInstance()
		{
			if (m_pDTE)
			{
				long nCount = 0;
				theApp.LoadCLR();
				HRESULT hr = S_FALSE;
				hr = m_pDTE->get_Windows(&m_pWindows);
				m_pWindows->AddRef();
				theApp.Lock();
				CComPtr<Events> pEvents;
				hr = m_pDTE->get_Events(&pEvents);
				CComQIPtr<Events2> pEvents2(pEvents);
				hr = pEvents2->get_DTEEvents(&m_pDTEEvents);
				hr = pEvents2->get_WindowEvents(0,&m_pWindowEvents);
				hr = pEvents2->get_ProjectsEvents(&m_pProjectsEvents);
				hr = pEvents2->get_ProjectItemsEvents(&m_pProjectItemsEvents);
				hr = ((CDTEEvents*)this)->DispEventAdvise((_WindowEvents*)m_pDTEEvents);
				hr = ((CDTEWindowEvents*)this)->DispEventAdvise((_WindowEvents*)m_pWindowEvents);
				hr = ((CDTEProjectsEvents*)this)->DispEventAdvise((_WindowEvents*)m_pProjectsEvents);
				hr = ((CDTEProjectItemsEvents*)this)->DispEventAdvise((_WindowEvents*)m_pProjectItemsEvents);
				theApp.Unlock();
				CComQIPtr<DTE2> pDTE2(m_pDTE);
				CComPtr<ToolWindows> pToolWindows;
				pDTE2->get_ToolWindows(&pToolWindows);
				CComPtr<OutputWindow> pOutWnd;
				pToolWindows->get_OutputWindow(&pOutWnd);
				if (pOutWnd)
				{
					CComPtr<OutputWindowPane> pPane;
					CComPtr<OutputWindowPanes> pPanes;
					pOutWnd->get_OutputWindowPanes(&pPanes);
					pPanes->Item(CComVariant(L"Tangram"), &pPane);
					if (pPane == nullptr)
					{
						pPanes->Add(CComBSTR(L"Tangram"), &pPane);
						pPane->OutputString(CComBSTR(L"Welcome to Tangram for Visual Studio!\r\n"));
					}
					m_pOutputWindowPane = pPane.Detach();
					m_pOutputWindowPane->AddRef();
				}
				if (::IsWindow(m_hHostWnd) == false)
				{
					m_hHostWnd = ::CreateWindowEx(WS_EX_PALETTEWINDOW, _T("Tangram Window Class"), _T("Tangram Designer"), WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, 0, 0, 0, 0, NULL, NULL, theApp.m_hInstance, NULL);
					m_hChildHostWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), _T(""), WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, m_hHostWnd, NULL, theApp.m_hInstance, NULL);
					theApp.m_mapValInfo[_T("hostwnd")] = CComVariant((long)m_hHostWnd);
					theApp.m_mapValInfo[_T("hostchildwnd")] = CComVariant((long)m_hChildHostWnd);
				}

				CComPtr<VxDTE::Windows> pWindows;
				pDTE2->get_Windows(&pWindows);
				CString strProperties = CString(vsWindowKindProperties);
				pWindows->Item(CComVariant(strProperties), &m_pPropertyWnd);
				m_pPropertyWnd->AddRef();
				hr = pEvents2->get_WindowVisibilityEvents(m_pPropertyWnd,&m_pPropertyWndVisibilityEvents);
				hr = ((CDTEWindowVisibilityEvents*)this)->DispEventAdvise((_WindowEvents*)m_pPropertyWndVisibilityEvents);
				//strProperties = CString(vsWindowKindToolbox);
				//pWindows->Item(CComVariant(strProperties), &m_pToolBoxWnd);
				//if (m_pToolBoxWnd)
				//{
				//	m_pToolBoxWnd->AddRef();
				//	long h = 0;
				//	m_pToolBoxWnd->get_HWnd(&h);
				//	if (h)
				//	{
				//		m_hHostWnd = (HWND)h;
				//		m_hChildHostWnd = ::GetWindow(m_hHostWnd, GW_CHILD);
				//	}
				//}

				//pWindows->get_Count(&nCount);
				//for (int i = 1; i <= nCount; i++)
				//{
				//	CComPtr<VxDTE::Window> pWnd;
				//	pWindows->Item(CComVariant(i), &pWnd);
				//	CComBSTR bstrName(L"");
				//	pWnd->get_Kind(&bstrName);
				//	CString str1 = OLE2T(bstrName);
				//	pWnd->get_ObjectKind(&bstrName);
				//	CString str2 = OLE2T(bstrName);
				//	long h = 0;
				//	pWnd->get_HWnd(&h);
				//	pWnd->get_Caption(&bstrName);
				//	CString str3 = OLE2T(bstrName);
				//}

				CComPtr<VxDTE::Documents> pDocuments;
				m_pDTE->get_Documents(&pDocuments);
				pDocuments->get_Count(&nCount);
				VARIANT_BOOL bVisible = false;
				for (int i = 1; i <= nCount; i++)
				{
					CComPtr<VxDTE::Document> pDoc;
					pDocuments->Item(CComVariant(i),&pDoc);
					CComBSTR bstrExtender(L"");
					CComPtr<VxDTE::Window> pWnd;
					pDoc->get_ActiveWindow(&pWnd);
					if (pWnd)
					{
						pWnd->get_Visible(&bVisible);
						if (bVisible)
						{
							AttachDoc(pDoc.p);
						}
					}
				}
			}
		}

		void CVSCloudAddin::OnWindowHiding(VxDTE::Window* Window)
		{ 
		}

		void CVSCloudAddin::OnWindowShowing(VxDTE::Window* pWnd) 
		{
			if (m_hPropertyForm == NULL)
			{
				HWND hwnd = NULL;
				CComPtr<VxDTE::Window> pLinkedWindowFrame;
				m_pPropertyWnd->get_LinkedWindowFrame(&pLinkedWindowFrame);
				if (pLinkedWindowFrame)
				{
					long h = 0;
					pLinkedWindowFrame->get_HWnd(&h);
					if (h)
					{
						hwnd = (HWND)h;
					}
					else
					{
						m_pPropertyWnd->Activate();
						hwnd = ::GetActiveWindow();
					}
					auto t = create_task([hwnd] {
						::Sleep(500);
						::PostMessage(hwnd, WM_INITPROPERTYGRID, 0, 0);
					});
				}
			}
		}

		//void CVSCloudAddin::OnBeforeClosing() 
		//{
		//}

		//void CVSCloudAddin::OnAfterClosing() 
		//{
		//}

		void CVSCloudAddin::OnItemAdded(VxDTE::Project* Project)
		{
			//CComBSTR bstrKind(L"");
			//Project->get_Kind(&bstrKind);
			//CString strKind = OLE2T(bstrKind);
			//if (strKind.CompareNoCase(m_strprjKindCSharpProject) == 0 || strKind.CompareNoCase(m_strprjKindVBProject) == 0)
			//{
			//	Project->get_FullName(&bstrKind);
			//	CString strPath = OLE2T(bstrKind);
			//	CTangramXmlParse* pParse = new CTangramXmlParse();
			//	int nPos = strPath.ReverseFind('\\');
			//	strPath = strPath.Left(nPos + 1) + _T("tangramresource.xml");
			//	strPath.MakeLower();
			//	if (!::PathFileExists(strPath))
			//	{
			//		CString strXml = _T("<Tangram><Forms></Forms></Tangram>");
			//		pParse->LoadXml(strXml);
			//		pParse->SaveFile(strPath);
			//	}
			//	else
			//	{
			//		pParse->LoadFile(strPath);
			//	}
			//	m_mapPrjXml[strPath] = pParse;
			//	nPos = strPath.ReverseFind('\\');
			//	CString strName = strPath.Mid(nPos + 1);
			//	CComPtr<VxDTE::ProjectItem> pItem;
			//	CComPtr<VxDTE::ProjectItems> pItems;
			//	Project->get_ProjectItems(&pItems);
			//	pItems->Item(CComVariant(strName), &pItem);
			//	if (!pItem)
			//	{
			//		HRESULT hr = pItems->AddFromFile(CComBSTR(strPath), &pItem);
			//		if (hr == S_OK)
			//		{
			//			CComBSTR bstrName(L"");
			//			pItem->get_Name(&bstrName);
			//			strName = OLE2T(bstrName);
			//			CComPtr<VxDTE::Properties> pProperties;
			//			pItem->get_Properties(&pProperties);
			//			if (pProperties)
			//			{
			//				CComPtr<VxDTE::Property> pProperty;
			//				pProperties->Item(CComVariant(L"ItemType"), &pProperty);
			//				if (pProperty)
			//				{
			//					pProperty->put_Value(CComVariant(L"EmbeddedResource"));
			//				}
			//			}
			//		}
			//	}
			//}
		}

		void CVSCloudAddin::OnItemRemoved(VxDTE::Project* Project) 
		{
			CComBSTR bstrKind(L"");
			Project->get_Kind(&bstrKind);
			CString strKind = OLE2T(bstrKind);
			if (strKind.CompareNoCase(m_strprjKindCSharpProject) == 0 || strKind.CompareNoCase(m_strprjKindVBProject) == 0)
			{
				Project->get_FullName(&bstrKind);
				CString strPath = OLE2T(bstrKind);
				int nPos = strPath.ReverseFind('\\');
				strPath = strPath.Left(nPos + 1) + _T("tangramresource.xml");
				strPath.MakeLower();
				auto it = m_mapPrjXml.find(strPath);
				if (it != m_mapPrjXml.end())
				{
					CTangramXmlParse* pParse = it->second;
					m_mapPrjXml.erase(it);
					pParse->Reflash(); 
					delete pParse;
				}
			}
		}

		void CVSCloudAddin::OnItemRenamed(VxDTE::Project* Project, BSTR OldName) 
		{
			CComBSTR bstrKind(L"");
			Project->get_Kind(&bstrKind);
			CString strKind = OLE2T(bstrKind);
			if (strKind.CompareNoCase(m_strprjKindCSharpProject) == 0 || strKind.CompareNoCase(m_strprjKindVBProject) == 0)
			{
			}
		}

		void CVSCloudAddin::OnItemAdded(ProjectItem* ProjectItem)
		{
		}

		void CVSCloudAddin::OnItemRemoved(ProjectItem* ProjectItem)
		{
		}

		void CVSCloudAddin::OnItemRenamed(ProjectItem* ProjectItem, BSTR OldName)
		{
		}

		//void CVSCloudAddin::OnWindowClosing(VxDTE::Window* Window)
		//{
		//}

		//void CVSCloudAddin::OnStartupComplete()
		//{
		//}

		void CVSCloudAddin::OnBeginShutdown()
		{
			HRESULT hr = S_FALSE;
			//hr = m_pDTE->get_Windows(&m_pWindows);
			//CString strProperties = CString(_T("{7a5a2942-c6a4-4d28-8c4d-7996c319206b}"));
			//VxDTE::Window* m_pToolBoxWnd = nullptr;
			//m_pWindows->Item(CComVariant(strProperties), &m_pToolBoxWnd);
			//if (m_pToolBoxWnd)
			//{
			//	CComPtr<VxDTE::Window> pLinkedWindowFrame;
			//	m_pToolBoxWnd->get_LinkedWindowFrame(&pLinkedWindowFrame);
			//	pLinkedWindowFrame->Close();
			//	//m_pToolBoxWnd->put_IsFloating(true);
			//	//m_pToolBoxWnd->Close(vsSaveChanges::vsSaveChangesYes);
			//}
			theApp.Lock();
			hr = ((CDTEEvents*)this)->DispEventUnadvise((_WindowEvents*)m_pDTEEvents);
			hr = ((CDTEWindowEvents*)this)->DispEventUnadvise((_WindowEvents*)m_pWindowEvents);
			hr = ((CDTEProjectsEvents*)this)->DispEventUnadvise((_WindowEvents*)m_pProjectsEvents);
			hr = ((CDTEProjectItemsEvents*)this)->DispEventUnadvise((_WindowEvents*)m_pProjectItemsEvents);
			hr = ((CDTEWindowVisibilityEvents*)this)->DispEventUnadvise((_WindowEvents*)m_pPropertyWndVisibilityEvents);
			theApp.Unlock();
		}

		void CVSCloudAddin::OnWindowCreated(VxDTE::Window* pWnd)
		{
			CComBSTR bstrKind(L"");
			pWnd->get_ObjectKind(&bstrKind);
			CString strKind = OLE2T(bstrKind);
			if (strKind.CompareNoCase(_T("7a5a2942-c6a4-4d28-8c4d-7996c319206b")) == 0)
			{
				TRACE(_T("Tangram Tool Window Created!\n"));
				long h = 0;
				pWnd->get_HWnd(&h);
				if (h)
				{
					m_hHostWnd = (HWND)h;
					m_hChildHostWnd = ::GetWindow(m_hHostWnd, GW_CHILD);
				}
			}
		}

		void CVSCloudAddin::OnWindowActivated(VxDTE::Window* _pWnd, VxDTE::Window* LostFocus)
		{
			CComPtr<VxDTE::Document> pDoc;
			m_pDTE->get_ActiveDocument(&pDoc);
			if (pDoc)
			{
				auto it = find(pDoc.p);
				if (it == end())
				{
					AttachDoc(pDoc.p);
				}
				else
				{
					m_pActiveDesignerWnd = it->second;
				}
			}
		}

		void CVSCloudAddin::ClearOutputPane()
		{
			if (m_pOutputWindowPane)
			{
				m_pOutputWindowPane->Clear();
			}
		}
	}
}
#endif

