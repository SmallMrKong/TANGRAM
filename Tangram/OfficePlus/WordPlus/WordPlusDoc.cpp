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
#include "../../WndNode.h"
#include "../../WndFrame.h"
#include "WordPlusDoc.h"
#include "WordAddin.h"
 
namespace OfficeCloudPlus
{
	namespace WordPlus
	{
		CWordDocument::CWordDocument(void)
		{
#ifdef _DEBUG
			theApp.m_nOfficeDocs++;
#endif			
			m_pDoc					= nullptr;
			m_pFrame				= nullptr;
			m_pPage					= nullptr;
			m_pTaskPaneFrame		= nullptr;
			m_pTaskPanePage			= nullptr;
			m_strTaskPaneXml		= _T("");
		}

		CWordDocument::~CWordDocument(void)
		{
#ifdef _DEBUG
			theApp.m_nOfficeDocs--;
#endif			
		}

		void CWordDocument::OnClose()
		{
			//this->DispEventUnadvise(m_pDoc);
		}

		void CWordDocument::InitDocument()
		{
			CWordCloudAddin* pAddin = (CWordCloudAddin*)theApp.m_pHostCore;
			IStream* pStream = 0;
			IStream* pStream2 = 0;
			HRESULT hr = ::CoMarshalInterThreadInterfaceInStream(IID_ITangram, theApp.m_pTangram, &pStream);
			hr = ::CoMarshalInterThreadInterfaceInStream(IID__Document, m_pDoc, &pStream2);
			auto t = create_task([pStream, pStream2, this]()
			{
				CoInitializeEx(NULL, COINIT_MULTITHREADED);
				ITangram* _pAddin = nullptr;
				HRESULT hr = ::CoGetInterfaceAndReleaseStream(pStream, IID_ITangram, (LPVOID *)&_pAddin);
				_Document* _pDoc = nullptr;
				hr = ::CoGetInterfaceAndReleaseStream(pStream2, IID__Document, (LPVOID *)&_pDoc);
				if (m_strDocXml == _T(""))
				{
					BSTR bstrXml(L"");
					_pAddin->GetDocXmlByKey(_pDoc, CComBSTR(L"Tangrams"), &bstrXml);
					m_strDocXml = OLE2T(bstrXml);
				}
				else
				{
					VARIANT_BOOL bRet = false;
					_pAddin->AddDocXml(_pDoc, CComBSTR(m_strDocXml), CComBSTR(L"Tangrams"), &bRet);
				}
				CoUninitialize();
			}).then([pAddin,this]()
			{
				CString _strXml = theApp.GetXmlData(_T("DocumentUI"), m_strDocXml);
				CString strKey = _T("<window>");
				CString strData = _T("");
				CString strVal = _T("");
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
					m_mapDocUIInfo[strData] = strVal;
				}

				_strXml = theApp.GetXmlData(_T("userforms"), m_strDocXml);
				strKey = _T("<window>");
				nPos = _strXml.Find(strKey);
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
				::PostMessage(pAddin->m_hHostWnd, WM_OPENDOCUMENT, (WPARAM)this, 0);
			});
		}

		CWordObject::CWordObject(void)
		{
			m_hForm				= NULL;
			m_hClient			= NULL;
			m_pWordPlusDoc		= nullptr;
			m_hTaskPane			= NULL;
			m_bDesignState		= false;
			m_bDesignTaskPane	= false;
		}

		void CWordObject::OnObjDestory()
		{
			if (m_pOfficeObj != nullptr)
			{
				CWordCloudAddin* pAddin = ((CWordCloudAddin*)theApp.m_pHostCore);

				if (pAddin->m_pActiveWordObject == this)
				{
					pAddin->m_pActiveWordObject = nullptr;
					theApp.m_pWndNode = nullptr;
				}

				auto it2 = m_pWordPlusDoc->find(m_hChildClient);
				if (it2 != m_pWordPlusDoc->end())
				{
					m_pWordPlusDoc->erase(it2);
				}

				size_t nCount = m_pWordPlusDoc->size();
				if (nCount > 0 && m_pWordPlusDoc->m_pFrame)
				{
					it2 = m_pWordPlusDoc->begin();
					m_pWordPlusDoc->m_pFrame->ModifyHost((long)::GetParent(it2->second->m_hChildClient));
					if (m_hTaskPane)
					{
						m_hTaskPane = NULL;
						m_pWordPlusDoc->m_pTaskPaneFrame->ModifyHost((long)it2->second->m_hTaskPaneChildWnd);
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
					if (m_pWordPlusDoc->m_pFrame)
					{
						::PostMessage(theApp.m_pHostCore->m_hHostWnd, WM_TANGRAMMSG, (WPARAM)m_pWordPlusDoc->m_pFrame->m_pWorkNode->m_pHostWnd->m_hWnd, 1992);
					}
					HRESULT hr = S_OK;
					hr = m_pWordPlusDoc->DispEventUnadvise(m_pWordPlusDoc->m_pDoc);
					m_pWordPlusDoc->erase(m_pWordPlusDoc->begin(), m_pWordPlusDoc->end());
					auto it2 = pAddin->find(m_pWordPlusDoc->m_pDoc);
					if (it2 != pAddin->end())
						pAddin->erase(it2);

					delete m_pWordPlusDoc;
				}
				auto it3 = pAddin->m_mapOfficeObjects2.find(m_hForm);
				if (it3 != pAddin->m_mapOfficeObjects2.end())
					pAddin->m_mapOfficeObjects2.erase(it3);
			}
		}
	}
}
