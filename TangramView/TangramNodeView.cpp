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
* http://www.cloudaddin.com
*
********************************************************************************/
// TangramNodeView.cpp : implementation file
//

#include "stdafx.h"
#include "def.h"
#include "dllmain.h"
#include "resource.h"
#include "TangramNodeView.h"
#include "TangramCategory.h"

// CTangramNodeView

IMPLEMENT_DYNCREATE(CTangramNodeView, CFormView)

CTangramNodeView::CTangramNodeView()
	: CFormView(IDD_TANGRAMNODEVIEW)
{
	m_nIndex = -1;
	m_bLayoutInit = false;
	m_bEnumActiveXID = false;
	m_bEnumComponentID = false;
	m_pNode = NULL;
	m_pWorkingNode = NULL;
	m_strCurAssembly = _T("");
	m_strCurAssemblys = _T("");
}

CTangramNodeView::~CTangramNodeView()
{
}

void CTangramNodeView::DoDataExchange(CDataExchange* pDX)
{
	CFormView::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TANGRAMCOMBO, m_ComBoTangram);
	DDX_Control(pDX, IDC_COMBO2, m_ComboStyle);
	DDX_Control(pDX, IDC_COMBONAMES, m_ComboNames);
	DDX_Control(pDX, IDC_COMBOROWS, m_ComboRows);
	DDX_Control(pDX, IDC_COMBOCOLS, m_ComboCols);
	DDX_Control(pDX, IDC_COMBOCATEGORY, m_ComboComponentCategory);
}

BEGIN_MESSAGE_MAP(CTangramNodeView, CFormView)
	ON_WM_MOUSEACTIVATE()
	ON_BN_CLICKED(IDC_CREATECLRCTRLBTN, &CTangramNodeView::OnCreateWndPageObj)
	ON_CBN_SELCHANGE(IDC_TANGRAMCOMBO, &CTangramNodeView::OnCbnSelchangeTangramCombo)
	ON_CBN_SELCHANGE(IDC_COMBONAMES, &CTangramNodeView::OnCbnSelchangeCombonames)
	ON_CBN_SELCHANGE(IDC_COMBOCATEGORY, &CTangramNodeView::OnCbnSelChangeComboCategory)
	ON_BN_CLICKED(IDC_BUTTON_SAVE, &CTangramNodeView::OnBnClickedButtonSave)
END_MESSAGE_MAP()


// CTangramNodeView diagnostics

#ifdef _DEBUG
void CTangramNodeView::AssertValid() const
{
	CFormView::AssertValid();
}

#ifndef _WIN32_WCE
void CTangramNodeView::Dump(CDumpContext& dc) const
{
	CFormView::Dump(dc);
}
#endif
#endif //_DEBUG


// CTangramNodeView message handlers


BOOL CTangramNodeView::Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext)
{
	// TODO: Add your specialized code here and/or call the base class

	return CFormView::Create(lpszClassName, lpszWindowName, dwStyle, rect, pParentWnd, nID, pContext);
}


void CTangramNodeView::OnInitialUpdate()
{
	CFormView::OnInitialUpdate();
	UpdateData(true);
	m_hWndCreate = ::GetDlgItem(m_hWnd, IDC_CREATECLRCTRLBTN);
	m_hRows = m_ComboRows.m_hWnd;
	m_hCols = m_ComboCols.m_hWnd;
	m_hRowsStatic = ::GetDlgItem(m_hWnd, IDC_ROWS);
	m_hColsStatic = ::GetDlgItem(m_hWnd, IDC_COLS);
	m_hCnnIDStatic = ::GetDlgItem(m_hWnd, IDC_STATICCNNID);
	m_hLayout = ::GetDlgItem(m_hWnd, IDC_LAYOUTS);
	m_hStyle = ::GetDlgItem(m_hWnd, IDC_STYLE);
	m_ComBoTangram.SetDroppedWidth(240);
	//::ShowWindow(m_hWndCreate, SW_HIDE);
	CString strVal = _T("");
	for (int i = 1; i <= 16; i++)
	{
		strVal.Format(_T("%d"), i);
		m_ComboRows.AddString(strVal);
		m_ComboCols.AddString(strVal);
	}
	m_ComboRows.SetCurSel(0);
	m_ComboCols.SetCurSel(0);
	m_ComboComponentCategory.AddString(_T("Layout"));
	m_ComboComponentCategory.AddString(_T("ActiveX Control"));
	m_ComboComponentCategory.AddString(_T(".NET Control"));
	m_ComboComponentCategory.AddString(_T("C++ Component"));
	m_ComboComponentCategory.AddString(_T("HostView"));
	m_ComboComponentCategory.SetCurSel(0);
	m_nIndex = 0;
	OnBnClickedCreateLayout();
	ModifyStyle(WS_CLIPCHILDREN, 0);
}


int CTangramNodeView::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message)
{
	if (m_pWorkingNode)
		m_pWorkingNode = NULL;
	return MA_ACTIVATE;
}

void CTangramNodeView::OnBnClickedPreviewbutton()
{
}

void CTangramNodeView::OnBnClickedHostview()
{
	IWndNode* pDesignerNode = NULL;
	theApp.m_pTangram->get_DesignNode(&pDesignerNode);
	if (pDesignerNode == NULL)
		return;
	pDesignerNode->put_Attribute(CComBSTR(L"id"), CComBSTR(L"hostview"));
	m_ComboComponentCategory.SetCurSel(0);
}

void CTangramNodeView::OnCreateWndPageObj()
{
	IWndNode* pDesignerNode = NULL;
	theApp.m_pTangram->get_DesignNode(&pDesignerNode);
	if (pDesignerNode == NULL)
		return;
	int nSel = -1;
	nSel = m_ComBoTangram.GetCurSel();
	long h = 0;
	HWND hWnd = NULL;
	HWND hPWnd = NULL;
	CString strID = _T("");
	CString strVal = _T("");
	m_ComBoTangram.GetLBText(nSel, strID);
	m_nIndex = m_ComboComponentCategory.GetCurSel();
	switch (m_nIndex)
	{
	case 0:
	case 3:
		{
			if (strID.CompareNoCase(_T("splitter")) == 0 || strID.CompareNoCase(_T("flat splitter")) == 0)
			{
				int nRow = 0;
				int nCol = 0;
				m_ComboRows.GetLBText(m_ComboRows.GetCurSel(), strVal);
				nRow = _wtoi(strVal);
				m_ComboCols.GetLBText(m_ComboCols.GetCurSel(), strVal);
				nCol = _wtoi(strVal);
				if (nRow*nCol>1)
				{
					CTangramXmlParse m_Parse;
					BOOL b = m_Parse.LoadXml(_T("<splitter><window></window></splitter>"));
					if (b)
					{
						CComBSTR bstrName(L"");
						pDesignerNode->get_Name(&bstrName);
						CTangramXmlParse* pWndNode = m_Parse.GetChild(_T("window"));
						RECT rc;
						::GetClientRect(m_hWnd, &rc);
						CComBSTR bstrNewName(L"");
						theApp.m_pTangram->GetNewLayoutNodeName(CComBSTR(L"Splitter"), pDesignerNode, &bstrNewName);
						CTangramXmlParse* pNewNode = pWndNode->AddSplitterNode(nRow, nCol, OLE2T(bstrNewName), rc.right, rc.bottom);
						CString str = m_Parse.xml();
						IWndNode* pParent = NULL;
						pDesignerNode->get_ParentNode(&pParent);
						if (pParent)
						{
							WndNodeType m_type;
							pParent->get_NodeType(&m_type);
							switch (m_type)
							{
							case WndNodeType::TNT_Splitter:
								pParent->get_Handle(&h);
								break;
							default:
								pDesignerNode->get_Handle(&h);
								break;
							}
						}
						else
						{
							pDesignerNode->get_Handle(&h);
						}
						hPWnd = (HWND)h;
						::SendMessage(hPWnd, WM_TANGRAMMSG, (WPARAM)pNewNode, (LPARAM)str.GetBuffer()/*pNode*/);
					}
				}
			}
			else
			{
				int nPages = 0;
				if (m_nIndex == 0)
				{
					m_ComboRows.GetLBText(m_ComboRows.GetCurSel(), strVal);
					nPages = _wtoi(strVal);
				}
				
				CString strObjID = _T("");
				m_ComboNames.GetLBText(m_ComboNames.GetCurSel(), strObjID);
				m_ComboStyle.GetLBText(m_ComboStyle.GetCurSel(), strVal);

				CTangramXmlParse m_Parse;
				BOOL b = m_Parse.LoadXml(_T("<splitter><window></window></splitter>"));
				if (b)
				{
					CComBSTR bstrName(L"");
					pDesignerNode->get_Name(&bstrName);
					CTangramXmlParse* pWndNode = m_Parse.GetChild(_T("window"));
					RECT rc;
					::GetClientRect(m_hWnd, &rc);
					CComBSTR bstrNewName(L"");
					theApp.m_pTangram->GetNewLayoutNodeName(CComBSTR(strObjID), pDesignerNode, &bstrNewName);
					CTangramXmlParse* pNewNode = pWndNode->AddTabNode(OLE2T(bstrNewName), strID, strObjID, strVal, nPages);
					CString str = m_Parse.xml();
					IWndNode* pParent = NULL;
					pDesignerNode->get_ParentNode(&pParent);

					WndNodeType m_type;
					if (pParent)
					{
						pParent->get_NodeType(&m_type);
						switch (m_type)
						{
						case WndNodeType::TNT_Splitter:
						{
							pParent->get_Handle(&h);
						}
						break;
						default:
							pDesignerNode->get_Handle(&h);
							break;
						}
					}
					else
					{
						pDesignerNode->get_Handle(&h);
					}

					hPWnd = (HWND)h;
					::SendMessage(hPWnd, WM_TANGRAMMSG, (WPARAM)pNewNode, (LPARAM)str.GetBuffer()/*pNode*/);
				}
			}
		}
		break;
	case 1:
		{
			if (strID != _T(""))
			{
				CString strObjID = strID;
				int nPos = strObjID.Find(_T("."));
				strObjID = strObjID.Mid(nPos + 1);
				nPos = strObjID.Find(_T("."));
				if (nPos != -1)
					strObjID = strObjID.Left(nPos);
				CComBSTR bstrNewName(L"");
				theApp.m_pTangram->GetNewLayoutNodeName(CComBSTR(strObjID), pDesignerNode, &bstrNewName);
				strObjID = OLE2T(bstrNewName);
				pDesignerNode->put_Attribute(CComBSTR(L"id"), CComBSTR(L"ActiveX"));
				if(strObjID!=_T(""))
					pDesignerNode->put_Attribute(CComBSTR(L"name"), CComBSTR(strObjID));
				pDesignerNode->put_Attribute(CComBSTR(L"cnnid"), CComBSTR(strID));
				pDesignerNode->get_Handle(&h);
				hWnd = (HWND)h;
				::SendMessage(hWnd, WM_TANGRAMMSG, (WPARAM)0, (LPARAM)strID.GetBuffer());
			}
		}
		break;
	case 2:
		{
			if (strID != _T(""))
			{
				strID += _T(",");
				strID += m_strCurAssembly;
				pDesignerNode->put_Attribute(CComBSTR(L"id"), CComBSTR(L"CLRCtrl"));
				pDesignerNode->put_Attribute(CComBSTR(L"cnnid"), CComBSTR(strID));
				pDesignerNode->get_Handle(&h);
				hWnd = (HWND)h;
				::SendMessage(hWnd, WM_TANGRAMMSG, (WPARAM)0, (LPARAM)strID.GetBuffer());
			}
		}
		break;
	case 4:
		OnBnClickedHostview();
		break;
	default:
		break;
	}
}

void CTangramNodeView::OnBnClickedCreatectrl()
{
	m_nIndex = 1;
	ReLayout(m_nIndex);
	for (int i = m_ComBoTangram.GetCount() - 1; i >= 0; i--)
	{
		m_ComBoTangram.DeleteString(i);
	}
	if (m_bEnumActiveXID == false)
	{
		CATID m_catid = CATID_Control;
		ULONG lFetched;
		CLSID ClsID;

		ICatInformation * pICatInformation = NULL;
		HRESULT hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr,
			NULL,
			CLSCTX_INPROC_SERVER,
			IID_ICatInformation,
			(void**)&pICatInformation);
		if (SUCCEEDED(hr))
		{
			IEnumGUID * pIEnumGUID = 0;
			HRESULT hr = pICatInformation->EnumClassesOfCategories(
				1,
				&m_catid,
				0,
				0,
				&pIEnumGUID);
			if (SUCCEEDED(hr))
			{
				pIEnumGUID->Reset();
				while (S_OK == pIEnumGUID->Next(1, &ClsID, &lFetched))
				{
					BSTR bstrID = ::SysAllocString(L"");
					hr = ::ProgIDFromCLSID(ClsID, &bstrID);
					if (S_OK == hr)
					{
						CString strID = OLE2T(bstrID);
						theApp.m_strCtrls += strID;
						theApp.m_strCtrls += _T(",");
						m_ComBoTangram.AddString(strID);
					}
				}
				pIEnumGUID->Release();
			}
			pICatInformation->Release();
			m_bEnumActiveXID = true;
		}
	}
	else
	{
		CString strLayouts = theApp.m_strCtrls;
		int nPos = strLayouts.Find(_T(","));
		while (nPos != -1)
		{
			CString strVal = strLayouts.Left(nPos);
			strLayouts = strLayouts.Mid(nPos + 1);
			nPos = strLayouts.Find(_T(","));
			m_ComBoTangram.AddString(strVal);
		}
	}
	m_ComBoTangram.SetCurSel(0);
	for (int i = m_ComboNames.GetCount() - 1; i >= 0; i--)
	{
		m_ComboNames.DeleteString(i);
	}
	for (int i = m_ComboStyle.GetCount() - 1; i >= 0; i--)
	{
		m_ComboStyle.DeleteString(i);
	}
	m_ComboNames.Invalidate(true);
	m_ComboStyle.Invalidate(true);
}

void CTangramNodeView::OnBnClickedCreateClrCtrl()
{
	m_nIndex = 2;
	ReLayout(m_nIndex);
	for (int i = m_ComBoTangram.GetCount() - 1; i >= 0; i--)
	{
		m_ComBoTangram.DeleteString(i);
	}
	if (m_strCurAssemblys != _T(""))
	{
		CString strLayouts = m_strCurAssemblys;
		int nPos = strLayouts.Find(_T(","));
		while (nPos != -1)
		{
			CString strVal = strLayouts.Left(nPos);
			strLayouts = strLayouts.Mid(nPos + 1);
			nPos = strLayouts.Find(_T(","));
			m_ComBoTangram.AddString(strVal);
		}
		m_ComBoTangram.SetCurSel(1);
		return;
	}

	m_ComBoTangram.AddString(_T("---Click me to Select .NET Control Library---"));
	m_ComBoTangram.SetCurSel(0);

	for (int i = m_ComboNames.GetCount() - 1; i >= 0; i--)
	{
		m_ComboNames.DeleteString(i);
	}
	for (int i = m_ComboStyle.GetCount() - 1; i >= 0; i--)
	{
		m_ComboStyle.DeleteString(i);
	}
	m_ComboNames.Invalidate(true);
	m_ComboStyle.Invalidate(true);
}

void CTangramNodeView::OnBnClickedCreateComponent()
{
	m_nIndex = 3;
	ReLayout(m_nIndex);
	for (int i = m_ComBoTangram.GetCount() - 1; i >= 0; i--)
	{
		m_ComBoTangram.DeleteString(i);
	}
	if (m_bEnumComponentID == false)
	{
		CATID m_catid = CATID_CloudUIComponentCategory;
		ULONG lFetched;
		CLSID ClsID;

		ICatInformation * pICatInformation = NULL;
		HRESULT hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr,
			NULL,
			CLSCTX_INPROC_SERVER,
			IID_ICatInformation,
			(void**)&pICatInformation);
		if (SUCCEEDED(hr))
		{
			IEnumGUID * pIEnumGUID = 0;
			HRESULT hr = pICatInformation->EnumClassesOfCategories(
				1,
				&m_catid,
				0,
				0,
				&pIEnumGUID);
			if (SUCCEEDED(hr))
			{
				pIEnumGUID->Reset();
				while (S_OK == pIEnumGUID->Next(1, &ClsID, &lFetched))
				{
					BSTR bstrID = ::SysAllocString(L"");
					hr = ::ProgIDFromCLSID(ClsID, &bstrID);
					if (S_OK == hr)
					{
						CString strID = OLE2T(bstrID);
						theApp.m_strComponents += strID;
						theApp.m_strComponents += _T(",");
						m_ComBoTangram.AddString(strID);
					}
				}
				pIEnumGUID->Release();
			}
			pICatInformation->Release();
			m_bEnumComponentID = true;
		}
	}
	else
	{
		CString strComponents = theApp.m_strComponents;
		int nPos = strComponents.Find(_T(","));
		while (nPos != -1)
		{
			CString strVal = strComponents.Left(nPos);
			strComponents = strComponents.Mid(nPos + 1);
			nPos = strComponents.Find(_T(","));
			m_ComBoTangram.AddString(strVal);
		}
	}
	m_ComBoTangram.SetCurSel(0);
	WPARAM wParam = MAKEWPARAM(IDC_TANGRAMCOMBO, CBN_SELCHANGE);
	::SendMessage(m_hWnd, WM_COMMAND, wParam, (LPARAM)m_ComBoTangram.m_hWnd);
	m_ComboNames.Invalidate(true);
	m_ComboStyle.Invalidate(true);
}


void CTangramNodeView::OnBnClickedCreateLayout()
{
	m_nIndex = 0;

	for (int i = m_ComBoTangram.GetCount() - 1; i >= 0; i--)
	{
		m_ComBoTangram.DeleteString(i);
	}
	for (int i = m_ComboNames.GetCount() - 1; i >= 0; i--)
	{
		m_ComboNames.DeleteString(i);
	}
	for (int i = m_ComboStyle.GetCount() - 1; i >= 0; i--)
	{
		m_ComboStyle.DeleteString(i);
	}
	if (m_bLayoutInit == false)
	{
		theApp.m_strLayouts += _T("Splitter,");
		m_ComBoTangram.AddString(_T("Splitter"));
		//theApp.m_strLayouts += _T("Flat Splitter,");
		//m_ComBoTangram.AddString(_T("Flat Splitter"));
		//theApp.m_strLayouts += _T("Tab,");
		//m_ComBoTangram.AddString(_T("Tab"));
		CATID m_catid = CATID_CloudUIContainerCategory;
		ULONG lFetched;
		CLSID ClsID;

		ICatInformation * pICatInformation = NULL;
		HRESULT hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr,
			NULL,
			CLSCTX_INPROC_SERVER,
			IID_ICatInformation,
			(void**)&pICatInformation);
		if (SUCCEEDED(hr))
		{
			IEnumGUID * pIEnumGUID = 0;
			HRESULT hr = pICatInformation->EnumClassesOfCategories(
				1,
				&m_catid,
				0,
				0,
				&pIEnumGUID);
			if (SUCCEEDED(hr))
			{
				pIEnumGUID->Reset();
				while (S_OK == pIEnumGUID->Next(1, &ClsID, &lFetched))
				{
					BSTR bstrID = ::SysAllocString(L"");
					hr = ::ProgIDFromCLSID(ClsID, &bstrID);
					if (S_OK == hr)
					{
						CString strID = OLE2T(bstrID);
						theApp.m_strLayouts += strID;
						theApp.m_strLayouts += _T(",");
						m_ComBoTangram.AddString(strID);
					}
				}
				pIEnumGUID->Release();
			}
			pICatInformation->Release();
			m_bLayoutInit = true;
		}
	}
	else
	{
		CString strLayouts = theApp.m_strLayouts;
		int nPos = strLayouts.Find(_T(","));
		while (nPos != -1)
		{
			CString strVal = strLayouts.Left(nPos);
			strLayouts = strLayouts.Mid(nPos + 1);
			nPos = strLayouts.Find(_T(","));
			m_ComBoTangram.AddString(strVal);
		}
	}
	m_ComBoTangram.SetCurSel(0);
	m_ComboNames.Invalidate(true);
	m_ComboStyle.Invalidate(true);
	this->ReLayout(m_nIndex);
}


void CTangramNodeView::OnCbnSelchangeTangramCombo()
{
	int nSel = m_ComBoTangram.GetCurSel();
	CString strID = _T("");
	m_ComBoTangram.GetLBText(nSel, strID);
	strID.MakeLower();
	switch (m_nIndex)
	{
	case 0:
	case 3:
		{
			CString strNames = _T("");
			map<CString, CString>::iterator it = theApp.m_mapName.find(strID);
			if (it == theApp.m_mapName.end())
			{
				if (strID.ReverseFind('.') != -1)
				{
					CComPtr<IDispatch> pDisp;
					pDisp.CoCreateInstance(CComBSTR(strID));
					if (pDisp)
					{
						CComQIPtr<ICreator> p(pDisp);
						if (p)
						{
							CComBSTR bstrNames(L"");
							p->get_Names(&bstrNames);
							strNames = OLE2T(bstrNames);
							strNames += _T(",");
							strNames.Replace(_T(",,"), _T(","));
							CString _strNames = strNames;
							int nPos = _strNames.Find(_T(","));
							CComBSTR bstrStyles(L"");
							while (nPos != -1)
							{
								CString strName = _strNames.Left(nPos);
								_strNames = _strNames.Mid(nPos + 1);
								nPos = _strNames.Find(_T(","));
								CString strKey = strID;
								strKey += _T("-");
								strKey += strName;
								{
									p->get_Tags(strName.AllocSysString(), &bstrStyles);
									CString strStyles = OLE2T(bstrStyles);
									strStyles += _T(",");
									strStyles.Replace(_T(",,"), _T(","));
									theApp.m_mapStyle[strKey] = strStyles;
								}
							}
						}
					}
				}
			}
			else
			{
				strNames = it->second;
			}
			if (strNames != _T(""))
			{
				for (int i = m_ComboNames.GetCount() - 1; i >= 0; i--)
				{
					m_ComboNames.DeleteString(i);
				}
				theApp.m_mapName[strID] = strNames;
				int nPos = strNames.Find(_T(","));
				while (nPos!=-1)
				{
					CString strName = strNames.Left(nPos);
					strNames = strNames.Mid(nPos + 1);
					nPos = strNames.Find(_T(","));
					m_ComboNames.AddString(strName);
				}
				m_ComboNames.SetCurSel(0);
				CString strName = _T("");
				m_ComboNames.GetLBText(0, strName);
				CString strKey = strID;
				strKey += _T("-");
				strKey += strName;
				auto it = theApp.m_mapStyle.find(strKey);
				if (it != theApp.m_mapStyle.end())
				{
					for (int i = m_ComboStyle.GetCount() - 1; i >= 0; i--)
					{
						m_ComboStyle.DeleteString(i);
					}
					CString strStyles = it->second;
					if (strStyles != _T(""))
					{
						int nPos = strStyles.Find(_T(","));
						while (nPos != -1)
						{
							CString strStyle = strStyles.Left(nPos);
							strStyles = strStyles.Mid(nPos + 1);
							nPos = strStyles.Find(_T(","));
							m_ComboStyle.AddString(strStyle);
						}
						m_ComboStyle.SetCurSel(0);
					}
				}
			}
			else
			{
				for (int i = m_ComboNames.GetCount() - 1; i >= 0; i--)
				{
					m_ComboNames.DeleteString(i);
				}
				for (int i = m_ComboStyle.GetCount() - 1; i >= 0; i--)
				{
					m_ComboStyle.DeleteString(i);
				}
			}
			ReLayout(m_nIndex);
		}
		break;
	case 1:
		break;
	case 2:
		{
			if (nSel == 0)
			{
				TCHAR szFile[MAX_PATH] = { 0 };
				::GetModuleFileName(NULL, szFile, MAX_PATH);
				CString strCfgFile = CString(szFile) + _T(".config");
				CTangramXmlParse m_Parse;
				if (m_Parse.LoadFile(strCfgFile))
				{
					CTangramXmlParse* pAssemblyBindingNode = m_Parse[_T("runtime")][_T("assemblyBinding")].GetChild(_T("probing"));
					CString	strVal = pAssemblyBindingNode->attr(_T("privatePath"), _T("")).MakeLower();
				}
				int nPos = strCfgFile.ReverseFind('\\');
				CString strAssemblyPath = strCfgFile.Left(nPos + 1);
				strAssemblyPath += _T("TangramAssemblies\\");
				BOOL isOpen = true;
				CString fileName = _T("");
				CString filter = _T("Assembly Files (*.dll)|*.dll||");
				CFileDialog openFileDlg(isOpen, strAssemblyPath, fileName, OFN_HIDEREADONLY | OFN_READONLY, filter, NULL);
				INT_PTR result = openFileDlg.DoModal();
				CString filePath = _T("");
				if (result == IDOK)
				{
					filePath = openFileDlg.GetPathName();
					int nPos = filePath.ReverseFind('\\');
					CString strLib = filePath.Mid(nPos + 1);
					nPos = strLib.Find(_T("."));
					strLib = strLib.Left(nPos).MakeLower();
					CString strAssemblys = _T("");
					if (m_strCurAssembly != strLib)
					{
						auto it = theApp.m_mapAssemblyInfo.find(strLib);
						if (it == theApp.m_mapAssemblyInfo.end())
						{
							BSTR bstrCtrls = ::SysAllocString(L"");
							theApp.m_pTangram->GetCLRControlString(CComBSTR(filePath), &bstrCtrls);
							CString strCtrls = OLE2T(bstrCtrls);
							::SysFreeString(bstrCtrls);
							if (strCtrls != _T(""))
							{
								m_strCurAssembly = strLib;
								CString strVal = _T("---Please Select .NET Control Library---");
								strAssemblys += strVal;
								strAssemblys += _T(",");
								strAssemblys += strCtrls;
								theApp.m_mapAssemblyInfo[strLib] = strAssemblys;
							}
							else
							{
								strAssemblys = _T("");
								::MessageBox(NULL, _T("Please Select a Valided .NET Control Library"), _T("Tangram Designer"), MB_ICONWARNING);
								return;
							}
						}
						else
						{
							strAssemblys = it->second;
						}
						m_strCurAssembly = strLib;
					}
					else
					{
						strAssemblys = m_strCurAssemblys;
					}
					if (strAssemblys != _T(""))
					{
						m_strCurAssemblys = strAssemblys;
						for (int i = m_ComBoTangram.GetCount() - 1; i >= 0; i--)
						{
							m_ComBoTangram.DeleteString(i);
						}
						nPos = strAssemblys.Find(_T(","));
						while (nPos != -1)
						{
							CString strID = strAssemblys.Left(nPos);
							strAssemblys = strAssemblys.Mid(nPos + 1);
							nPos = strAssemblys.Find(_T(","));
							m_ComBoTangram.AddString(strID);
						}
						m_ComBoTangram.SetCurSel(1);
					}
				}
			}
		}
		break;
	default:
		break;
	}
	m_ComboNames.Invalidate(true);
	m_ComboStyle.Invalidate(true);
}


void CTangramNodeView::OnCbnSelchangeCombonames()
{
	int nSel = m_ComBoTangram.GetCurSel();
	CString strID = _T("");
	CString strName = _T("");
	m_ComBoTangram.GetLBText(nSel, strID);
	strID.MakeLower();
	nSel = m_ComboNames.GetCurSel();
	m_ComboNames.GetLBText(nSel, strName);
	CString strKey = strID;
	strKey += _T("-");
	strKey += strName;
	switch (m_nIndex)
	{
	case 0:
	{
		auto it = theApp.m_mapStyle.find(strKey);
		if (it != theApp.m_mapStyle.end())
		{
			for (int i = m_ComboStyle.GetCount() - 1; i >= 0; i--)
			{
				m_ComboStyle.DeleteString(i);
			}
			CString strStyles = it->second;
			if (strStyles != _T(""))
			{
				int nPos = strStyles.Find(_T(","));
				while (nPos != -1)
				{
					CString strStyle = strStyles.Left(nPos);
					strStyles = strStyles.Mid(nPos + 1);
					nPos = strStyles.Find(_T(","));
					m_ComboStyle.AddString(strStyle);
				}
				m_ComboStyle.SetCurSel(0);
			}
		}
	}
		break;
	default:
		break;
	}
}

void CTangramNodeView::ReLayout(int nIndex)
{
	RECT rect;
	int nSel = m_ComBoTangram.GetCurSel();
	CString strID = _T("");
	if(nSel>=0)
		m_ComBoTangram.GetLBText(nSel, strID);
	strID.MakeLower();
	switch (nIndex)
	{
	case 0:
	case 3:
		{
			if (strID.CompareNoCase(_T("splitter")) == 0 || strID.CompareNoCase(_T("flat splitter")) == 0)
			{
				m_ComBoTangram.GetWindowRect(&rect);
				ScreenToClient(&rect);
				//::SetWindowPos(m_ComBoTangram.m_hWnd, HWND_TOP, rect.left, rect.top, 80, rect.bottom - rect.top, SWP_NOACTIVATE);
				::ShowWindow(m_hRowsStatic, SW_SHOW);
				::SetWindowText(m_hRowsStatic, _T("Rows:"));
				::ShowWindow(m_hRows, SW_SHOW);
				::ShowWindow(m_hCols, SW_SHOW);
				::ShowWindow(m_ComboNames.m_hWnd, SW_HIDE);
				::ShowWindow(m_ComboStyle.m_hWnd, SW_HIDE);
				::ShowWindow(m_hLayout, SW_HIDE);
				::ShowWindow(m_hStyle, SW_HIDE);
				::ShowWindow(m_hColsStatic, SW_SHOW);
				//::SetWindowPos(m_hRowsStatic, HWND_TOP, rect.left + 100, rect.top + 3, 30, rect.bottom - rect.top, SWP_NOACTIVATE);
				//::SetWindowPos(m_hRows, HWND_TOP, rect.left + 140, rect.top, 40, rect.bottom - rect.top, SWP_NOACTIVATE);
				//::SetWindowPos(m_hColsStatic, HWND_TOP, rect.left + 200, rect.top + 3, 30, rect.bottom - rect.top, SWP_NOACTIVATE);
				//::SetWindowPos(m_hCols, HWND_TOP, rect.left + 230, rect.top, 40, rect.bottom - rect.top, SWP_NOACTIVATE);
				//::SetWindowPos(m_hWndCreate, HWND_TOP, 20 + 70, rect.bottom + 10, 60, rect.bottom - rect.top, SWP_NOACTIVATE);
			}
			else
			{
				m_ComBoTangram.GetWindowRect(&rect);
				ScreenToClient(&rect);
				//::SetWindowPos(m_ComBoTangram.m_hWnd, HWND_TOP, rect.left, rect.top, 240, rect.bottom - rect.top, SWP_NOACTIVATE);
				m_ComboNames.GetWindowRect(&rect);
				ScreenToClient(&rect);
				if (m_nIndex == 0)
				{
					//::SetWindowPos(m_ComboNames.m_hWnd, HWND_TOP, rect.left, rect.top, 150, rect.bottom - rect.top, SWP_NOACTIVATE);
					//::ShowWindow(m_hRowsStatic, SW_SHOW);
					::SetWindowText(m_hRowsStatic, _T("Pages:"));
					::ShowWindow(m_hRows, SW_SHOW);
					//::SetWindowPos(m_hRowsStatic, HWND_TOP, rect.left + 160, rect.top + 2, 35, rect.bottom - rect.top, SWP_NOACTIVATE);
					//::SetWindowPos(m_hRows, HWND_TOP, rect.left + 200, rect.top, 40, rect.bottom - rect.top, SWP_NOACTIVATE);
				}
				else
				{
					//::SetWindowPos(m_ComboNames.m_hWnd, HWND_TOP, rect.left, rect.top, 240, rect.bottom - rect.top, SWP_NOACTIVATE);
					::ShowWindow(m_hRowsStatic, SW_HIDE);
					::ShowWindow(m_hRows, SW_HIDE);
				}
				::ShowWindow(m_hLayout, SW_SHOW);
				::ShowWindow(m_hStyle, SW_SHOW);
				::ShowWindow(m_ComboNames.m_hWnd, SW_SHOW);
				::ShowWindow(m_ComboStyle.m_hWnd, SW_SHOW);
				::ShowWindow(m_hCols, SW_HIDE);
				::ShowWindow(m_hColsStatic, SW_HIDE);
				//m_ComboStyle.GetWindowRect(&rect);
				//ScreenToClient(&rect);
				//::SetWindowPos(m_ComboStyle.m_hWnd, HWND_TOP, rect.left, rect.top, 240, rect.bottom - rect.top, SWP_NOACTIVATE);
				//::SetWindowPos(m_hWndCreate, HWND_TOP, 20 + 70, rect.bottom + 10, 60, rect.bottom - rect.top, SWP_NOACTIVATE);
			}
		}
		break;
	case 1:
	case 2:
		{
			m_ComBoTangram.GetWindowRect(&rect);
			ScreenToClient(&rect);
			//::SetWindowPos(m_ComBoTangram.m_hWnd, HWND_TOP, rect.left, rect.top, 240, rect.bottom - rect.top, SWP_NOACTIVATE);
			m_ComboNames.ShowWindow(SW_HIDE);
			m_ComboStyle.ShowWindow(SW_HIDE);
			::ShowWindow(m_hRowsStatic, SW_HIDE);
			::ShowWindow(m_hLayout, SW_HIDE);
			::ShowWindow(m_hStyle, SW_HIDE);
			::ShowWindow(m_hRows, SW_HIDE);
			::ShowWindow(m_hCols, SW_HIDE);
			::ShowWindow(m_hColsStatic, SW_HIDE);
			//::SetWindowPos(m_hWndCreate, HWND_TOP, 20 + 70, rect.bottom + 10, 60, rect.bottom - rect.top, SWP_NOACTIVATE);
		}
		break;
	default:
		break;
	}
	this->Invalidate(true);
}


void CTangramNodeView::OnCbnSelChangeComboCategory()
{
	m_nIndex = m_ComboComponentCategory.GetCurSel();
	switch (m_nIndex)
	{
	case 0:
	{
		OnBnClickedCreateLayout();
	}
	break;
	case 1:
		OnBnClickedCreatectrl();
		break;
	case 2:
		OnBnClickedCreateClrCtrl();
		break;
	case 3:
		OnBnClickedCreateComponent();
		break;
	case 4:
		//OnBnClickedHostview();
		break;
	default:
		break;
	}
}

void CTangramNodeView::OnBnClickedButtonSave()
{
	IWndNode* pDesignerNode = NULL;
	theApp.m_pTangram->get_DesignNode(&pDesignerNode);
	if (pDesignerNode == NULL)
		return;
	IWndNode* pRootNode = NULL;
	pDesignerNode->get_RootNode(&pRootNode);
	if (pRootNode)
	{
		IWndFrame* pFrame = NULL;
		pRootNode->get_Frame(&pFrame);
		if (pFrame)
		{
			long nHandle = 0;
			pFrame->get_HWND(&nHandle);
			HWND hWnd = (HWND)nHandle;
			if (::IsWindow(hWnd))
			{
				::SendMessage(hWnd, WM_TANGRAMSAVE, 1965, 1965);
			}
		}
	}
}
