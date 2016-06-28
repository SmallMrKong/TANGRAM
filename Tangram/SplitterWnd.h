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


#pragma once

// CSplitterNodeWnd

class CSplitterNodeWnd : public CSplitterWnd,
	public CComObjectRootBase,
	public IDispatchImpl<IWndContainer, &__uuidof(IWndContainer), &LIBID_Tangram,  1,  0>
{
	DECLARE_DYNCREATE_ATL(CSplitterNodeWnd)
public:

	BEGIN_COM_MAP(CSplitterNodeWnd)
		COM_INTERFACE_ENTRY(IWndContainer)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()
	void Lock(){}
	void Unlock(){}

	STDMETHOD(Save)();

#ifdef _DEBUG
	virtual void AssertValid() const;
#ifndef _WIN32_WCE
	virtual void Dump(CDumpContext& dc) const;
#endif
#endif

	int m_Vmin,m_Vmax,m_Hmin,m_Hmax;
	COLORREF		rgbLeftTop;
	COLORREF		rgbMiddle;
	COLORREF		rgbRightBottom;
protected:
	CSplitterNodeWnd();           // protected constructor used by dynamic creation
	virtual ~CSplitterNodeWnd();

	BOOL			m_bCreated;
	BOOL			m_bFlatSplitter;
	
	vector<RECT>	ColRect;
	vector<RECT>	RowRect;

	bool			bColMoving ;
	bool			bRowMoving ;

	CWndNode*		m_pWndNode;
	CCreateContext* m_pContext;

	ULONG InternalAddRef(){return 1;}
	ULONG InternalRelease(){return 1;}

	BOOL PreCreateWindow(CREATESTRUCT& cs);
	BOOL OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult);
	BOOL Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext = NULL);
	void RecalcLayout();
	void StartTracking(int ht);
	void StopTracking(BOOL bAccept);
	void OnDrawSplitter(CDC* pDC, ESplitType nType, const CRect& rect);
	void PostNcDestroy();
	void DrawAllSplitBars(CDC* pDC, int cxInside, int cyInside);
	CWnd* GetActivePane(int* pRow = NULL, int* pCol = NULL);

	afx_msg void OnPaint();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnDestroy();
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	afx_msg LRESULT OnActivePage(WPARAM wParam,LPARAM lParam);
	afx_msg LRESULT OnSplitterNodeAdd(WPARAM wParam,LPARAM lParam);
	afx_msg LRESULT OnActiveTangramObj(WPARAM wParam,LPARAM lParam);
	afx_msg LRESULT OnGetTangramObj(WPARAM wParam,LPARAM lParam);
	DECLARE_MESSAGE_MAP()
};
