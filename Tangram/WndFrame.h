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
class ATL_NO_VTABLE CWndFrame : 
	public CComObjectRootBase,	
	public CWindowImpl<CWndFrame, CWindow>,
	public IDispatchImpl<IWndFrame, &IID_IWndFrame, &LIBID_Tangram,  1,  0>
{
public:
	CWndFrame();           
	virtual ~CWndFrame();           

	BOOL									m_bDetached;
	BOOL									m_bOuterInited;
	BOOL									m_bDesignerState;
	BOOL									m_bNeedOuterChanged;

	int										m_x;
	int										m_y;
	LONG									m_left;
	LONG									m_top;
	LONG									m_right;
	LONG									m_bottom;
	LONG									m_left2;
	LONG									m_top2;
	LONG									m_right2;
	LONG									m_bottom2;
	HWND									m_hHostWnd;

	HWND									m_hLWnd;
	HWND									m_hTWnd;
	HWND									m_hRWnd;
	HWND									m_hBWnd;
	HWND									m_hOuterLWnd;
	HWND									m_hOuterTWnd;
	HWND									m_hOuterRWnd;
	HWND									m_hOuterBWnd;

	CString									m_strFrameName;	
	CString									m_strCurrentKey;

	CWndPage*								m_pPage;
	CWndNode*								m_pHostNode;
	CWndNode*								m_pWorkNode;
	CWndNode*								m_pBindingNode;
	CBKWnd*									m_pBKWnd;
	map<CString, CWndNode*>					m_mapNode;
	map<CString, VARIANT>					m_mapVal;
	CComObject<CWndNodeCollection>*			m_pRootNodes;

	void Lock(){}
	void Unlock(){}
	void Destroy();
	void UpdateWndNode();
	void UpdateDesignerTreeInfo();
	BOOL CreateWndPage();
	CWndNode* OpenXtmlDocument(CString strKey, CString	strFile);

	STDMETHOD(get_HWND)(long* pVal);
	STDMETHOD(get_WndPage)(IWndPage** pVal);
	STDMETHOD(get_CurrentNavigateKey)(BSTR* pVal);
	STDMETHOD(get_VisibleNode)(IWndNode** pVal);
	STDMETHOD(get_RootNodes)(IWndNodeCollection** pNodeColletion);
	STDMETHOD(get_FrameData)(BSTR bstrKey, VARIANT* pVal);
	STDMETHOD(put_FrameData)(BSTR bstrKey, VARIANT newVal);
	STDMETHOD(get_DesignerState)(VARIANT_BOOL* pVal);
	STDMETHOD(put_DesignerState)(VARIANT_BOOL newVal);

	STDMETHOD(Attach)(void);
	STDMETHOD(Detach)(void);
	STDMETHOD(ModifyHost)(long hHostWnd);
	STDMETHOD(Extend)(BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode);
	STDMETHOD(GetXml)(BSTR bstrRootName, BSTR* bstrRet);

	void OnFinalMessage(HWND hWnd);

	BEGIN_COM_MAP(CWndFrame)
		COM_INTERFACE_ENTRY(IWndFrame)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()

	BEGIN_MSG_MAP(CWndFrame)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_HSCROLL, OnHScroll)
		MESSAGE_HANDLER(WM_VSCROLL, OnHScroll)
		MESSAGE_HANDLER(WM_NCDESTROY, OnNcDestroy)
		MESSAGE_HANDLER(WM_SETHELPWND, OnSetHelpWnd)
		MESSAGE_HANDLER(WM_SETHELPWNDEX, OnSetHelpWndEx)
		MESSAGE_HANDLER(WM_PARENTNOTIFY, OnParentNotify)
		MESSAGE_HANDLER(WM_TANGRAMMSG, OnTangramMsg)
		MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
		MESSAGE_HANDLER(WM_TANGRAMDATA, OnGetMe)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGING, OnWindowPosChanging)
	END_MSG_MAP()

protected:
	ULONG InternalAddRef(){ return 1; }
	ULONG InternalRelease(){ return 1; }

private:
	LRESULT OnGetMe(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnHScroll(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnNcDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnSetHelpWnd(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnTangramMsg(UINT, WPARAM, LPARAM, BOOL&);
	//LRESULT OnEraseBkgnd(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnSetHelpWndEx(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnParentNotify(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnMouseActivate(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnWindowPosChanging(UINT, WPARAM, LPARAM, BOOL&);
};
