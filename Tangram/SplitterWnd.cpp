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
*
********************************************************************************/

// TangramSplitterWnd.cpp : implementation file
//

#include "stdafx.h"
#include "CloudAddinApp.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "NodeWnd.h"
#include "SplitterWnd.h"
#include "TangramHtmlTreeWnd.h"
#include "VisualStudioPlus\CloudAddinVSAddin.h"

struct AUX_DATA
{
	// system metrics
	int cxVScroll, cyHScroll;
	int cxIcon, cyIcon;

	int cxBorder2, cyBorder2;

	// device metrics for screen
	int cxPixelsPerInch, cyPixelsPerInch;

	// convenient system color
	HBRUSH hbrWindowFrame;
	HBRUSH hbrBtnFace;

	// color values of system colors used for CToolBar
	COLORREF clrBtnFace, clrBtnShadow, clrBtnHilite;
	COLORREF clrBtnText, clrWindowFrame;

	// standard cursors
	HCURSOR hcurWait;
	HCURSOR hcurArrow;
	HCURSOR hcurHelp;       // cursor used in Shift+F1 help

	// special GDI objects allocated on demand
	HFONT   hStatusFont;
	HFONT   hToolTipsFont;
	HBITMAP hbmMenuDot;

// Implementation
	AUX_DATA();
	~AUX_DATA();
	void UpdateSysColors();
	void UpdateSysMetrics();
};

extern AFX_DATA AUX_DATA afxData;
#ifndef AFX_CX_BORDER
#define AFX_CX_BORDER CX_BORDER 
#endif

#ifndef AFX_CY_BORDER
#define AFX_CY_BORDER CY_BORDER 
#endif
#define CX_BORDER   1
#define CY_BORDER   1

// CSplitterNodeWnd
enum HitTestValue
{
	noHit                   = 0,
	vSplitterBox            = 1,
	hSplitterBox            = 2,
	bothSplitterBox         = 3,        // just for keyboard
	vSplitterBar1           = 101,
	vSplitterBar15          = 115,
	hSplitterBar1           = 201,
	hSplitterBar15          = 215,
	splitterIntersection1   = 301,
	splitterIntersection225 = 525
};

IMPLEMENT_DYNCREATE_ATL(CSplitterNodeWnd, CSplitterWnd)

CSplitterNodeWnd::CSplitterNodeWnd()
{
	m_bCreated		= true;
	m_bFlatSplitter = true;
	m_Vmin = m_Vmax = m_Hmin = m_Hmax = 0;
}

CSplitterNodeWnd::~CSplitterNodeWnd()
{
}

BEGIN_MESSAGE_MAP(CSplitterNodeWnd, CSplitterWnd)
	ON_WM_SIZE()
	ON_WM_PAINT()
	ON_WM_DESTROY()
	ON_WM_MOUSEMOVE()
	ON_WM_MOUSEACTIVATE()
	ON_MESSAGE(WM_TABCHANGE,OnActivePage)
	ON_MESSAGE(WM_TANGRAMGETNODE, OnGetTangramObj)
	ON_MESSAGE(WM_TANGRAMMSG, OnSplitterNodeAdd)
	ON_MESSAGE(WM_TGM_SETACTIVEPAGE,OnActiveTangramObj)
END_MESSAGE_MAP()

// CSplitterNodeWnd diagnostics

#ifdef _DEBUG
void CSplitterNodeWnd::AssertValid() const
{
	//CSplitterWnd::AssertValid();
}

#ifndef _WIN32_WCE
void CSplitterNodeWnd::Dump(CDumpContext& dc) const
{
	CSplitterWnd::Dump(dc);
}
#endif
#endif //_DEBUG

BOOL CSplitterNodeWnd::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	return true;
}

LRESULT CSplitterNodeWnd::OnSplitterNodeAdd(WPARAM wParam, LPARAM lParam)
{
	if (lParam == 1992)
	{
		return 0;
	}
	if (wParam == 0x01000)
		return 0;
	CTangramXmlParse* pNewNXmlode = (CTangramXmlParse*)wParam;
	if (pNewNXmlode == nullptr)
	{
		return 0;
	}
	CString strID = pNewNXmlode->attr(_T("id"), _T(""));
	IWndNode* _pNode = nullptr;
	CString str = (LPCTSTR)lParam;
	CWndNode* pOldNode = (CWndNode*)theApp.m_pDesignWindowNode;
	CTangramXmlParse* pOld = pOldNode->m_pHostParse;

	long	m_nRow;
	theApp.m_pDesignWindowNode->get_Row(&m_nRow);
	long	m_nCol;
	theApp.m_pDesignWindowNode->get_Col(&m_nCol);
	IWndNode* _pOldNode = nullptr;
	m_pWndNode->GetNode(m_nRow, m_nCol, &_pOldNode);
	CWndNode* _pOldNode2 = (CWndNode*)_pOldNode;
	CTangramXmlParse* _pOldParse = _pOldNode2->m_pHostParse;
	m_pWndNode->ExtendEx(m_nRow, m_nCol, CComBSTR(L""), str.AllocSysString(), &_pNode);
	CWndNode* pNode = (CWndNode*)_pNode;
	if (pNode&&pOldNode)
	{
		CTangramNodeVector::iterator it;
		it = find(m_pWndNode->m_vChildNodes.begin(), m_pWndNode->m_vChildNodes.end(), pOldNode);

		if (it != m_pWndNode->m_vChildNodes.end())
		{
			pNode->m_pRootObj = m_pWndNode->m_pRootObj;
			pNode->m_pParentObj = m_pWndNode;
			m_pWndNode->m_pRootObj->m_mapLayoutNodes[((CWndNode*)pNode)->m_strName] = (CWndNode*)pNode;
			CTangramNodeVector vec = pNode->m_vChildNodes;
			CWndNode* pChildNode = nullptr;
			for (auto it2 : pNode->m_vChildNodes)
			{
				pChildNode = it2;
				pChildNode->m_pRootObj = m_pWndNode->m_pRootObj;
			}
			CTangramXmlParse* pNew = pNode->m_pHostParse;
			CTangramXmlParse* pRet = m_pWndNode->m_pHostParse->ReplaceNode(pOld, pNew, _T(""));
			m_pWndNode->m_vChildNodes.erase(it);
			m_pWndNode->m_vChildNodes.push_back(pNode);
			pOldNode->m_pHostWnd->DestroyWindow();
			CString strXml = m_pWndNode->m_pRootObj->m_pTangramParse->xml();
			theApp.m_pHostDesignUINode = m_pWndNode->m_pFrame->m_pWorkNode;
			if (theApp.m_pHostDesignUINode)
			{
				HWND hFrame = theApp.m_pHostDesignUINode->m_pFrame->m_hWnd;
				::SendMessage(hFrame, WM_TANGRAMGETDESIGNWND, 0, 1);//Modify Form Designer for Visual Studio
				CTangramHtmlTreeWnd* pTreeCtrl = theApp.m_pDocDOMTree;
				pTreeCtrl->DeleteItem(theApp.m_pDocDOMTree->m_hFirstRoot);

				if (pTreeCtrl->m_pHostXmlParse)
				{
					delete pTreeCtrl->m_pHostXmlParse;
				}
				theApp.InitDesignerTreeCtrl(strXml);
				theApp.m_pHostDesignUINode->m_pDocXmlParseNode = pTreeCtrl->m_pHostXmlParse;
			}
#ifndef _WIN64
			if (theApp.m_strExeName == _T("devenv"))
			{
				VSCloudPlus::VisualStudioPlus::CVSCloudAddin* pAddin = (VSCloudPlus::VisualStudioPlus::CVSCloudAddin*)theApp.m_pHostCore;
				if(pAddin->m_pOutputWindowPane)
					pAddin->m_pOutputWindowPane->OutputString(strXml.AllocSysString());
			}
#endif
			this->RecalcLayout();
		}
	}

	return -1;
}

LRESULT CSplitterNodeWnd::OnActiveTangramObj(WPARAM wParam, LPARAM lParam)
{
	theApp.HostPosChanged(m_pWndNode);
	return -1;
}

LRESULT CSplitterNodeWnd::OnActivePage(WPARAM wParam, LPARAM lParam)
{
	if(m_pWndNode)
	{
		//CComPtr<IWndNodeCollection> pCol;
		//m_pWndNode->get_ChildNodes(&pCol);
		//long nCount = 0;
		//pCol->get_NodeCount(&nCount);
		//for(int i = 0;i< nCount;i++)
		//{
		//	CComPtr<IWndNode> pNode;
		//	pCol->get_Item(i,&pNode);
		//	HWND hWnd;
		//	pNode->get_Handle((LONGLONG*)&hWnd);
		//	::SendMessage(hWnd,WM_TABCHANGE,wParam,lParam);
		//}
	}
	if (theApp.m_pDocDOMTree&&theApp.m_pCloudAddinCLRProxy)
	{
		//CWndFrame* pFrame = m_pWndNode->m_pRootObj->m_pFrame;
		//if (pFrame->m_bDesignerState)
		theApp.m_pCloudAddinCLRProxy->SelectNode(m_pWndNode);
	}
	return CWnd::DefWindowProc(WM_TABCHANGE,wParam,lParam);;
}

void CSplitterNodeWnd::StartTracking(int ht)
{
	ASSERT_VALID(this);
	if (ht == noHit)
		return;

	//theApp.m_bSplitterTracking = true;
	CWndNode* pNode = m_pWndNode->m_pFrame->m_pWorkNode;
	if(pNode->m_pHostClientView)
	{
		pNode->m_pHostWnd->ModifyStyle(WS_CLIPSIBLINGS,0);
	}

	GetInsideRect(m_rectLimit);

	if (ht >= splitterIntersection1 && ht <= splitterIntersection225)
	{
		// split two directions (two tracking rectangles)
		int row = (ht - splitterIntersection1) / 15;
		int col = (ht - splitterIntersection1) % 15;

		GetHitRect(row + vSplitterBar1, m_rectTracker);
		int yTrackOffset = m_ptTrackOffset.y;
		m_bTracking2 = true;
		GetHitRect(col + hSplitterBar1, m_rectTracker2);
		m_ptTrackOffset.y = yTrackOffset;
	}
	else if (ht == bothSplitterBox)
	{
		// hit on splitter boxes (for keyboard)
		GetHitRect(vSplitterBox, m_rectTracker);
		int yTrackOffset = m_ptTrackOffset.y;
		m_bTracking2 = true;
		GetHitRect(hSplitterBox, m_rectTracker2);
		m_ptTrackOffset.y = yTrackOffset;

		// center it
		m_rectTracker.OffsetRect(0, m_rectLimit.Height()/2);
		m_rectTracker2.OffsetRect(m_rectLimit.Width()/2, 0);
	}
	else
	{
		// only hit one bar
		GetHitRect(ht, m_rectTracker);
	}

	// steal focus and capture
	SetCapture();
	SetFocus();

	// make sure no updates are pending
	RedrawWindow(NULL, NULL, RDW_ALLCHILDREN | RDW_UPDATENOW);

	// set tracking state and appropriate cursor
	m_bTracking = true;
	OnInvertTracker(m_rectTracker);
	if (m_bTracking2)
		OnInvertTracker(m_rectTracker2);
	m_htTrack = ht;
}

void CSplitterNodeWnd::StopTracking(BOOL bAccept)
{
	if (!m_bTracking)
		return;
	CWndFrame* pFrame = m_pWndNode->m_pFrame;
	CWndNode* pNode = pFrame->m_pWorkNode;
	if (pNode->m_pHostClientView)
	{
		pNode->m_pHostWnd->ModifyStyle(0,WS_CLIPSIBLINGS);
		::InvalidateRect(pFrame->m_hWnd, NULL, false); 
		pNode->m_pHostWnd->Invalidate();
	}

	CSplitterWnd::StopTracking(bAccept);

	RecalcLayout();
	if(bAccept)
	{
		theApp.HostPosChanged(m_pWndNode);
	}
}

void CSplitterNodeWnd::PostNcDestroy()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	CSplitterWnd::PostNcDestroy();
	delete m_pWndNode;
	delete this;
}

void CSplitterNodeWnd::RecalcLayout()
{
	if (GetDlgItem(IdFromRowCol(0, 0)) == NULL)
		return;
	if(m_nMaxCols >= 2)
	{
		int LimitWidth = 0;
		int LimitWndCount = 0;
		int Width = 0;
		RECT SplitterWndRect;
 		GetWindowRect(&SplitterWndRect);
		Width = SplitterWndRect.right - SplitterWndRect.left - m_nMaxCols * m_cxSplitterGap - LimitWidth + m_cxBorder * m_nMaxCols;
		if(Width < 0)
			return;
	}	
	
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	CSplitterWnd::RecalcLayout();
	//CTangramNodeVector::iterator it2;
	//CWndNode* pChildNode = NULL;
	//for (it2 = m_pWndNode->m_vChildNodes.begin(); it2 != m_pWndNode->m_vChildNodes.end(); it2++)
	//{
	//	pChildNode = *it2;
	//	CWnd* pWnd = pChildNode->m_pHostWnd;
	//	int nRow = pChildNode->m_nRow;
	//	int nCol = pChildNode->m_nCol;
	//	pWnd->Invalidate();
	//	//this->GetRowInfo()
	//}
}

BOOL CSplitterNodeWnd::Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext)
{
	m_pWndNode = theApp.m_pWndNode;
	m_pWndNode -> m_pHostWnd = this;
	m_pWndNode -> m_nViewType = Splitter;
	m_pWndNode -> m_nID = nID;
	m_pWndNode -> m_pDisp = this;
	m_pWndNode -> m_pRootObj->m_mapLayoutNodes[m_pWndNode->m_strName] = m_pWndNode;
	m_pContext = pContext;
	int r,g,b;
	CComBSTR bstrVal(L"");
	m_pWndNode->get_Attribute(CComBSTR(L"lefttopcolor"),&bstrVal);
	if(!CString(bstrVal).IsEmpty())
	{
		_stscanf_s(CString(bstrVal),_T("RGB(%d,%d,%d)"),&r,&g,&b);
		rgbLeftTop = RGB(r,g,b);
	}
	else
	{
		rgbLeftTop = RGB(240,240,240);
	}

	bstrVal.Empty();
	m_pWndNode->get_Attribute(CComBSTR(L"rightbottomcolor"),&bstrVal);
	if(!CString(bstrVal).IsEmpty())
	{
		_stscanf_s(CString(bstrVal),_T("RGB(%d,%d,%d)"),&r,&g,&b);
		rgbRightBottom = RGB(r,g,b);
	}
	else
	{
		rgbRightBottom = RGB(240,240,240);//RGB(71,104,158)
	}
	bstrVal.Empty();
	m_pWndNode->get_Attribute(CComBSTR(L"middlecolor"),&bstrVal);
	if(!CString(bstrVal).IsEmpty())
	{
		_stscanf_s(CString(bstrVal),_T("RGB(%d,%d,%d)"),&r,&g,&b);
		rgbMiddle = RGB(r,g,b);
	}
	else
	{
		rgbMiddle = RGB(240,240,240);
	}
	bstrVal.Empty();
	m_pWndNode->get_Attribute(CComBSTR(L"splitterwidth"),&bstrVal);
	if(!CString(bstrVal).IsEmpty())
	{
		m_cxSplitterGap = m_cySplitterGap = m_cxSplitter =  m_cySplitter = _ttoi(CString(bstrVal)); 

	}
	else
	{
		m_cxSplitterGap = m_cySplitterGap = m_cxSplitter =  m_cySplitter = 7; 
	}
	bstrVal.Empty();
	m_pWndNode->get_Attribute(CComBSTR(L"borderwidth"),&bstrVal);
	if(!CString(bstrVal).IsEmpty())
	{
		m_cxBorder = m_cyBorder = _ttoi(CString(bstrVal));
	}
	else
	{
		m_cxBorder = m_cyBorder = 2;
	}
	bstrVal.Empty();
	m_pWndNode->get_Attribute(CComBSTR(L"vmin"),&bstrVal);
	if(!CString(bstrVal).IsEmpty())
		m_Vmin = _ttoi(CString(bstrVal));
	else
		m_Vmin = 0;

	bstrVal.Empty();
	m_pWndNode->get_Attribute(CComBSTR(L"vmax"),&bstrVal);
	if(!CString(bstrVal).IsEmpty())
	{
		m_Vmax = _ttoi(CString(bstrVal));
		if(m_Vmax <= 0)
			m_Vmax = 65535;
	}
	else
		m_Vmax = 65535;
	bstrVal.Empty();
	m_pWndNode->get_Attribute(CComBSTR(L"hmin"),&bstrVal);
	if(!CString(bstrVal).IsEmpty())
		m_Hmin = _ttoi(CString(bstrVal));
	else
		m_Hmin = 0;

	bstrVal.Empty();
	m_pWndNode->get_Attribute(CComBSTR(L"hmax"),&bstrVal);
	if(!CString(bstrVal).IsEmpty())
	{
		m_Hmax = _ttoi(CString(bstrVal));
		if(m_Hmax <= 0)
			m_Hmax = 65535;
	}
	else
		m_Hmax = 65535;

	if(m_pWndNode->m_nStyle)
		m_bFlatSplitter=false;

	m_pWndNode->m_nRows = m_pWndNode->m_pHostParse->attrInt(TGM_ROWS,0);
	m_pWndNode->m_nCols = m_pWndNode->m_pHostParse->attrInt(TGM_COLS,0);

	bColMoving = 0;
	bRowMoving = 0;

	if (nID == 0)
		nID = 1;

	if(CreateStatic(pParentWnd,m_pWndNode->m_nRows,m_pWndNode->m_nCols,dwStyle,nID))
	{
		CString strWidth = m_pWndNode->m_pHostParse->attr(TGM_WIDTH,_T(""));
		strWidth += _T(",");
		CString strHeight = m_pWndNode->m_pHostParse->attr(TGM_HEIGHT,_T(""));
		strHeight += _T(",");

		int nWidth,nHeight,nPos;
		CString strW,strH, strOldWidth,strName;

		strOldWidth = strWidth;
		long nSize = m_pWndNode->m_pHostParse->GetCount();
		int nIndex = 0;
		CTangramXmlParse* pSubItem = m_pWndNode->m_pHostParse->GetChild(nIndex);
		for(int i = 0; i<m_pWndNode->m_nRows; i++)
		{	
			nPos = strHeight.Find(_T(","));
			strH = strHeight.Left(nPos);
			strHeight = strHeight.Mid(nPos+1);
			nHeight = _ttoi(strH);

			strWidth = strOldWidth;
			for(int j = 0; j<m_pWndNode->m_nCols; j++)
			{
				nPos = strWidth.Find(_T(","));
				strW = strWidth.Left(nPos);
				strWidth = strWidth.Mid(nPos+1);
				nWidth = _ttoi(strW);
				strName = pSubItem->attr(TGM_NAME,_T(""));
				strName.Trim();
				strName.MakeLower();
				CWndNode* pObj = new CComObject<CWndNode>;
				pObj -> m_pRootObj = m_pWndNode -> m_pRootObj;
				if (nIndex<nSize)
					pObj->m_pHostParse = pSubItem;

				pObj->m_pParentObj = m_pWndNode;
				pObj->m_pFrame = m_pWndNode->m_pFrame;

				m_pWndNode->AddChildNode(pObj);
				theApp.InitWndNode(pObj);

				if(pObj->m_pObjClsInfo)
				{
					pObj->m_nRow = i;
					pObj->m_nCol = j;

					pObj->m_nWidth = nWidth;
					pObj->m_nHeigh = nHeight;
					if (pContext->m_pNewViewClass == nullptr)
						pContext->m_pNewViewClass = RUNTIME_CLASS(CNodeWnd);
					CreateView(pObj->m_nRow,pObj->m_nCol,pObj->m_pObjClsInfo,CSize(pObj->m_nWidth,pObj->m_nHeigh),pContext);
				}
				nIndex++;
				if (nIndex<nSize)
					pSubItem = m_pWndNode->m_pHostParse->GetChild(nIndex);
			}
		}
		SetWindowPos(NULL,rect.left,rect.top,rect.right-rect.left,rect.bottom-rect.top, SWP_NOZORDER | SWP_NOREDRAW);
		if (m_pWndNode->m_pPage)
			m_pWndNode->m_pPage->Fire_NodeCreated(m_pWndNode);

		SetWindowText(m_pWndNode->m_strWebObjName);
		m_bCreated = true;
		return true;
	}
	return false;
}

LRESULT CSplitterNodeWnd::OnGetTangramObj(WPARAM wParam, LPARAM lParam)
{
	if(m_pWndNode)
		return (LRESULT)m_pWndNode;
	return (long)0;
}

void CSplitterNodeWnd::OnPaint()
{
	if( m_bFlatSplitter )
	{
		ASSERT_VALID(this);
		CPaintDC dc(this);

		CRect rectClient;
		GetClientRect(&rectClient);

		CRect rectInside;
		GetInsideRect(rectInside);
		rectInside.InflateRect(4,4);


		// draw the splitter boxes
		if (m_bHasVScroll && m_nRows < m_nMaxRows)
		{
			OnDrawSplitter(&dc, splitBox, CRect(rectInside.right, rectClient.top,rectClient.right, rectClient.top + m_cySplitter));
		}

		if (m_bHasHScroll && m_nCols < m_nMaxCols)
		{
			OnDrawSplitter(&dc, splitBox,
				CRect(rectClient.left, rectInside.bottom,
				rectClient.left + m_cxSplitter, rectClient.bottom));
		}

		// extend split bars to window border (past margins)
		DrawAllSplitBars(&dc, rectInside.right, rectInside.bottom);
		// draw splitter intersections (inside only)
		GetInsideRect(rectInside);
		dc.IntersectClipRect(rectInside);
		CRect rect;
		rect.top = rectInside.top;
		for (int row = 0; row < m_nRows - 1; row++)
		{
			rect.top += m_pRowInfo[row].nCurSize + m_cyBorderShare;
			rect.bottom = rect.top + m_cySplitter;
			rect.left = rectInside.left;
			for (int col = 0; col < m_nCols - 1; col++)
			{
				rect.left += m_pColInfo[col].nCurSize + m_cxBorderShare;
				rect.right = rect.left + m_cxSplitter;
				OnDrawSplitter(&dc, splitIntersection, rect);
				rect.left = rect.right + m_cxBorderShare;
			}
			rect.top = rect.bottom + m_cxBorderShare;
		}
	}
	else 
	{
		CSplitterWnd::OnPaint();
	}
}

void CSplitterNodeWnd::OnDrawSplitter(CDC* pDC, ESplitType nType, const CRect& rectArg)
{
	if( m_bFlatSplitter )
	{
		if (pDC == nullptr)
		{
			RedrawWindow(rectArg, NULL, RDW_INVALIDATE|RDW_NOCHILDREN);
			return;
		}
		ASSERT_VALID(pDC);
	  ;

		// otherwise, actually draw
		CRect rect = rectArg;
		switch (nType)
		{
		case splitBorder:
			//ASSERT(afxData.bWin4);
			pDC->Draw3dRect(rect, rgbLeftTop, rgbRightBottom);
			rect.InflateRect(-AFX_CX_BORDER, -AFX_CY_BORDER);
			pDC->Draw3dRect(rect, rgbLeftTop, rgbRightBottom);
			
			return;

		case splitIntersection:
			
			//ASSERT(afxData.bWin4);
			break;

		case splitBox:
			//if (afxData.bWin4)
			{
				pDC->Draw3dRect(rect,  rgbLeftTop, rgbRightBottom);
				rect.InflateRect(-AFX_CX_BORDER, -AFX_CY_BORDER);
				pDC->Draw3dRect(rect, rgbLeftTop, rgbRightBottom);
				rect.InflateRect(-AFX_CX_BORDER, -AFX_CY_BORDER);
				break;
			}
			// fall through...
		case splitBar:
			
			{
				if((rect.bottom - rect.top) > (rect.right - rect.left))
				{
					pDC->FillSolidRect(rect, rgbMiddle);	
				}
				else
				{
					pDC->FillSolidRect(rect, rgbMiddle);
				}
			}
			break;

		default:
			ASSERT(false);  // unknown splitter type
		}
	}

	else 
		CSplitterWnd::OnDrawSplitter(pDC, nType, rectArg);
}

BOOL CSplitterNodeWnd::PreCreateWindow(CREATESTRUCT& cs)
{
	cs.lpszClass = theApp.m_lpszSplitterClass;
	cs.style |= WS_CLIPSIBLINGS;
	return CSplitterWnd::PreCreateWindow(cs);
}

void CSplitterNodeWnd::DrawAllSplitBars(CDC* pDC, int cxInside, int cyInside)
{
	ColRect.clear();
	RowRect.clear();
	ASSERT_VALID(this);

	int col;
	int row;
	CRect rect;

	// draw pane borders
	GetClientRect(rect);
	int x = rect.left;
	for (col = 0; col < m_nCols; col++)
	{
		int cx = m_pColInfo[col].nCurSize + 2*m_cxBorder;
		if (col == m_nCols-1 && m_bHasVScroll)
			cx += afxData.cxVScroll - CX_BORDER;
		int y = rect.top;
		for (row = 0; row < m_nRows; row++)
		{
			int cy = m_pRowInfo[row].nCurSize + 2*m_cyBorder;
			if (row == m_nRows-1 && m_bHasHScroll)
				cy += afxData.cyHScroll - CX_BORDER;
			OnDrawSplitter(pDC, splitBorder, CRect(x, y, x+cx, y+cy));
			y += cy + m_cySplitterGap - 2*m_cyBorder;
		}
		x += cx + m_cxSplitterGap - 2*m_cxBorder;
	}


	// draw column split bars
	GetClientRect(rect);
	rect.left += m_cxBorder;
	for (col = 0; col < m_nCols - 1; col++)
	{
		rect.left += m_pColInfo[col].nCurSize + m_cxBorderShare;
		rect.right = rect.left + m_cxSplitter;
		if (rect.left > cxInside)
			break;      // stop if not fully visible
		//ColumnsplitBar = rect;
		ColRect.push_back(rect);
		OnDrawSplitter(pDC, splitBar, rect);
		 
		rect.left = rect.right + m_cxBorderShare;
	}

	// draw row split bars
	GetClientRect(rect);
	rect.top += m_cyBorder;
	for (row = 0; row < m_nRows - 1; row++)
	{
		rect.top += m_pRowInfo[row].nCurSize + m_cyBorderShare;
		rect.bottom = rect.top + m_cySplitter;
		if (rect.top > cyInside)
			break;      // stop if not fully visible
		//RowsplitBar = rect;
		RowRect.push_back(rect);
		OnDrawSplitter(pDC, splitBar, rect);
		
		rect.top = rect.bottom + m_cyBorderShare;
	}
}

CWnd* CSplitterNodeWnd::GetActivePane(int* pRow , int* pCol)
{
	CWnd* pView = nullptr;
	pView = GetFocus();

	// make sure the pane is a child pane of the splitter
	if (pView != nullptr && !IsChildPane(pView, pRow, pCol))
		pView = nullptr;


	return pView;
}

int CSplitterNodeWnd::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message)
{
	if(m_pWndNode->m_pParentObj)
	{
		if (m_pWndNode->m_pParentObj->m_nViewType & TabbedWnd)
			m_pWndNode->m_pParentObj->m_pVisibleXMLObj = m_pWndNode;
	}
	if (theApp.m_pDocDOMTree)
	{
		theApp.m_pCloudAddinCLRProxy->SelectNode(m_pWndNode);
	}
	return MA_ACTIVATE;// CSplitterWnd::OnMouseActivate(pDesktopWnd, nHitTest, message);
}

STDMETHODIMP CSplitterNodeWnd::Save()
{
	CString strWidth = _T("");
	CString strHeight = _T("");

	int minCx,minCy;

	for(int i = 0; i< m_pWndNode->m_nRows; i++)
	{
		int iHeight;
		CString strH;
		GetRowInfo(i,iHeight,minCy);
		strH.Format(_T("%d,"),iHeight);

		strHeight+=strH;
	}

	for(int j = 0; j < m_pWndNode -> m_nCols; j++)
	{
		int iWidth;
		CString strW;
		GetColumnInfo(j,iWidth,minCx);
		strW.Format(_T("%d,"),iWidth);

		strWidth += strW;
	}

	m_pWndNode->put_Attribute(CComBSTR(TGM_HEIGHT),CComBSTR(strHeight));
	m_pWndNode->put_Attribute(CComBSTR(TGM_WIDTH),CComBSTR(strWidth));
	
	return S_OK;
}

void CSplitterNodeWnd::OnMouseMove(UINT nFlags, CPoint point)
{
	if(point.x < m_Hmin && point.y < m_Vmin)   
	{   
		CSplitterWnd::OnMouseMove(nFlags, CPoint(m_Hmin,m_Vmin)); 
	}  

	else if(point.x < m_Hmin && point.y > m_Vmin && point.y < m_Vmax)			
	{   
		CSplitterWnd::OnMouseMove(nFlags, CPoint(m_Hmin,point.y)); 
	} 

	else if(point.x < m_Hmin && point.y > m_Vmax)   
	{   
		CSplitterWnd::OnMouseMove(nFlags,  CPoint(m_Hmin,m_Vmax)); 
	} 

	else if(point.x > m_Hmax && point.y < m_Vmin)   
	{   
		CSplitterWnd::OnMouseMove(nFlags, CPoint(m_Hmax,m_Vmin)); 
	}  

	else if(point.x > m_Hmax && point.y > m_Vmin && point.y < m_Vmax)			
	{   
		CSplitterWnd::OnMouseMove(nFlags, CPoint(m_Hmax,point.y)); 
	} 

	else if(point.x > m_Hmax && point.y > m_Vmax)   
	{   
		CSplitterWnd::OnMouseMove(nFlags,  CPoint(m_Hmax,m_Vmax)); 
	} 

	else if(point.x > m_Hmin && point.x < m_Hmax && point.y > m_Vmax)   
	{   
		CSplitterWnd::OnMouseMove(nFlags,  CPoint(point.x,m_Vmax)); 
	} 
	else if(point.x > m_Hmin && point.x < m_Hmax && point.y < m_Vmin)   
	{   
		CSplitterWnd::OnMouseMove(nFlags,  CPoint(point.x,m_Vmin)); 
	} 
	else   
	{  
		CSplitterWnd::OnMouseMove(nFlags,   point);   
	}   

	CDC *pDC = GetDC();
	for (int col = 0; col < m_nCols - 1; col++)
	{
		if(PtInRect( &ColRect.at(col),point) && bColMoving == 0)
		{
			pDC->FillSolidRect(&(ColRect.at(col)),rgbMiddle);
		}
	}

	for (int row = 0; row < m_nRows - 1; row++)
	{
		if(PtInRect( &RowRect.at(row),point))
		{
			pDC->FillSolidRect(&(RowRect.at(row)), rgbMiddle);
		}
	}
}

void CSplitterNodeWnd::OnSize(UINT nType, int cx, int cy)
{
	__super::OnSize(nType, cx, cy);
	if(m_pColInfo != nullptr)
		RecalcLayout();
	Invalidate(false);
}

void CSplitterNodeWnd::OnDestroy()
{
	m_pWndNode->Fire_Destroy();
	__super::OnDestroy();
}
