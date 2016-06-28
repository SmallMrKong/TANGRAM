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


#pragma once

class CTangramHtmlTreeWnd;

class CNodeWnd : public CWnd
{
	DECLARE_DYNCREATE(CNodeWnd)
public:
	BOOL					m_bBKWnd;
	BOOL					m_bEraseBkgnd;
	BOOL					m_bCreateExternal;

	CWndNode*				m_pWndNode;
	CWndNode*				m_pParentNode;
	CTangramHtmlTreeWnd*	m_pXHtmlTree;

	IOleInPlaceActiveObject* m_pOleInPlaceActiveObject;

	BOOL PreTranslateMessage(MSG* pMsg);
	BOOL Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext = NULL);
#ifdef _DEBUG
	virtual void AssertValid() const;
#ifndef _WIN32_WCE
	virtual void Dump(CDumpContext& dc) const;
#endif
#endif

protected:
	CNodeWnd();           // protected constructor used by dynamic creation
	virtual ~CNodeWnd();
	void PostNcDestroy();
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	afx_msg void OnDestroy();
	afx_msg void OnWindowPosChanged(WINDOWPOS* lpwndpos);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg LRESULT OnActiveTangramObj(WPARAM wParam,LPARAM lParam);
	afx_msg LRESULT OnTangramMsg(WPARAM wParam,LPARAM lParam);
	afx_msg LRESULT OnTabChange(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnGetTangramObj(WPARAM wParam, LPARAM lParam);
	DECLARE_MESSAGE_MAP()
};

class CBKWnd : public CWindowImpl<CBKWnd, CWindow>
{
public:
	CBKWnd(void);
	CWndNode* m_pWndNode;

	BEGIN_MSG_MAP(CBKWnd)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGING, OnWindowPosChanged)
	END_MSG_MAP()

private:
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&){ return 0; };
	LRESULT OnMouseActivate(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnWindowPosChanged(UINT, WPARAM, LPARAM, BOOL&);
	void OnFinalMessage(HWND hWnd);
};
