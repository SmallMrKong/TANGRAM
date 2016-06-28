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
* http://www.cloudaddin.com
*
********************************************************************************/

#pragma once

// CTabPageWnd

class CTabPageWnd : public CWnd
{
public:
	CTabPageWnd();
	virtual ~CTabPageWnd();

protected:
	virtual void PostNcDestroy();
	DECLARE_MESSAGE_MAP()
};

// CTangramTabCtrlWnd

class CTangramTabCtrlWnd : public CMFCTabCtrl
{
	DECLARE_DYNAMIC(CTangramTabCtrlWnd)

public:
	CTangramTabCtrlWnd();
	virtual ~CTangramTabCtrlWnd();
	virtual BOOL SetActiveTab(int iTab);

	int m_nCurSelTab;
public:
	IWndNode*	m_pWndNode;

protected:
	DECLARE_MESSAGE_MAP()
	afx_msg LRESULT OnCreatePage(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnActivePage(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnModifyPage(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnTgmSetCaption(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnActiveTangramObj(WPARAM wParam, LPARAM lParam);
	virtual void PostNcDestroy();
};


