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

#include "tangram.h"
#include "afxwin.h"
// CTangramNodeView form view

class CTangramNodeView : public CFormView
{
	DECLARE_DYNCREATE(CTangramNodeView)

protected:
	CTangramNodeView();           // protected constructor used by dynamic creation
	virtual ~CTangramNodeView();

public:
	IWndNode* m_pNode;
	IWndNode* m_pWorkingNode;
	int m_nIndex;
	BOOL m_bEnumActiveXID;
	BOOL m_bEnumComponentID;
	BOOL m_bLayoutInit;
	HWND m_hWndCreate;

	HWND m_hRows;
	HWND m_hCols;
	HWND m_hRowsStatic;
	HWND m_hColsStatic;
	HWND m_hCnnIDStatic;
	HWND m_hLayout;
	HWND m_hStyle;

	CString m_strCurAssembly;
	CString m_strCurAssemblys;
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_TANGRAMNODEVIEW };
#endif
#ifdef _DEBUG
	virtual void AssertValid() const;
#ifndef _WIN32_WCE
	virtual void Dump(CDumpContext& dc) const;
#endif
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	void ReLayout(int nIndex);
	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext = NULL);
	virtual void OnInitialUpdate();
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	afx_msg void OnBnClickedPreviewbutton();
	afx_msg void OnCreateWndPageObj();
	afx_msg void OnBnClickedCreatectrl();
	afx_msg void OnBnClickedCreateClrCtrl();
	afx_msg void OnBnClickedCreateComponent();
	afx_msg void OnBnClickedCreateLayout();
	afx_msg void OnCbnSelchangeTangramCombo();
	CComboBox m_ComBoTangram;
	CComboBox m_ComboStyle;
	CComboBox m_ComboNames;
	afx_msg void OnCbnSelchangeCombonames();
	CComboBox m_ComboRows;
	CComboBox m_ComboCols;
	afx_msg void OnBnClickedHostview();
	CComboBox m_ComboComponentCategory;
	afx_msg void OnCbnSelChangeComboCategory();
	afx_msg void OnBnClickedButtonSave();
};


