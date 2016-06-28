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
//
// TangramToolWindow.h
//

#pragma once

#include <initguid.h>
#include "..\TangramPackageUI\Resource.h"
#include "TangramVSIApp.h"

// {624ed9c3-bdfd-41fa-96c3-7c824ea32e3d}
DEFINE_GUID(EnvironmentColorsCategory, 
0x624ed9c3, 0xbdfd, 0x41fa, 0x96, 0xc3, 0x7c, 0x82, 0x4e, 0xa3, 0x2e, 0x3d);

#define VS_RGBA_TO_COLORREF(rgba) (rgba & 0x00FFFFFF)

class CAppXmlToolWnd;
class CAppXmlWndPane :
	public CComObjectRootEx<CComSingleThreadModel>,
	public VsWindowPaneFromResource<CAppXmlWndPane, IDD_TangramPackage_DLG>,
	public VsWindowFrameEventSink<CAppXmlWndPane>,
	public VSL::ISupportErrorInfoImpl<
		InterfaceSupportsErrorInfoList<IVsWindowPane,
		InterfaceSupportsErrorInfoList<IVsWindowFrameNotify,
		InterfaceSupportsErrorInfoList<IVsWindowFrameNotify3> > > >,
		public IVsBroadcastMessageEvents
{
	VSL_DECLARE_NOT_COPYABLE(CAppXmlWndPane)

protected:

	// Protected constructor called by CComObject<CAppXmlWndPane>::CreateInstance.
	CAppXmlWndPane() :
		VsWindowPaneFromResource(),
		m_hBackground(nullptr),
		m_BroadcastCookie(VSCOOKIE_NIL)
	{
		m_pAppXmlToolWnd = nullptr;
	}

	~CAppXmlWndPane() 
	{
	}
	
public:
	CAppXmlToolWnd* m_pAppXmlToolWnd;
BEGIN_COM_MAP(CAppXmlWndPane)
	COM_INTERFACE_ENTRY(IVsWindowPane)
	COM_INTERFACE_ENTRY(IVsWindowFrameNotify)
	COM_INTERFACE_ENTRY(IVsWindowFrameNotify3)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IVsBroadcastMessageEvents)
END_COM_MAP()

BEGIN_MSG_MAP(CAppXmlWndPane)
	COMMAND_HANDLER(IDC_CLICKME_BTN, BN_CLICKED, OnButtonClick)
	MESSAGE_HANDLER(WM_CTLCOLORDLG, OnCtlColorDlg)
END_MSG_MAP()

	// Function called by VsWindowPaneFromResource at the end of SetSite; at this point the
	// window pane is constructed and sited and can be used, so this is where we can initialize
	// the event sink by siting it.
	void PostSited(IVsPackageEnums::SetSiteResult /*result*/)
	{
		VsWindowFrameEventSink<CAppXmlWndPane>::SetSite(GetVsSiteCache());
		CComPtr<IVsShell> spShell = GetVsSiteCache().GetCachedService<IVsShell, SID_SVsShell>();
		spShell->AdviseBroadcastMessages(this, &m_BroadcastCookie);
		InitVSColors();
		if (theApp.m_pDTE == NULL)
		{
			CHKHR(GetVsSiteCache().QueryService(SID_SDTE, &theApp.m_pDTE));
			if (theApp.m_pDTE)
				theApp.InitInstance();
		}
	}

	// Function called by VsWindowPaneFromResource at the end of ClosePane.
	// Perform necessary cleanup.
	void PostClosed();

	//	CComPtr<IVsShell> spShell = GetVsSiteCache().GetCachedService<IVsShell, SID_SVsShell>();
	//	if (nullptr != spShell && VSCOOKIE_NIL != m_BroadcastCookie)
	//	{
	//		spShell->UnadviseBroadcastMessages(m_BroadcastCookie);
	//		m_BroadcastCookie = VSCOOKIE_NIL;
	//	}
	//}

	// Callback function called by ToolWindowBase when the size of the window changes.
	void OnFrameSize(int x, int y, int w, int h)
	{
		// Center button.
		CWindow button(this->GetDlgItem(IDC_CLICKME_BTN));
		RECT buttonRectangle;
		button.GetWindowRect(&buttonRectangle);

		OffsetRect(&buttonRectangle, -buttonRectangle.left, -buttonRectangle.top);

		int iLeft = (w - buttonRectangle.right) / 2; 
		if (iLeft <= 0)
		{
			iLeft = 5;
		}

		int iTop = (h - buttonRectangle.bottom) / 2; 
		if (iTop <= 0)
		{
			iTop = 5;
		}

		button.SetWindowPos(NULL, iLeft, iTop, 0, 0, SWP_NOSIZE);
	}

	LRESULT OnButtonClick(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
	{
		// Load the message from the resources.
		CComBSTR strMessage;
		VSL_CHECKBOOL_GLE(strMessage.LoadStringW(_AtlBaseModule.GetResourceInstance(), IDS_BUTTONCLICK_MESSAGE));

		// Get the title of the message box (it is the same as the tool window's title).
		CComBSTR strTitle;
		VSL_CHECKBOOL_GLE(strTitle.LoadStringW(_AtlBaseModule.GetResourceInstance(), IDS_WINDOW_TITLE));

		// Get the UI Shell service.
		CComPtr<IVsUIShell> spIVsUIShell = GetVsSiteCache().GetCachedService<IVsUIShell, SID_SVsUIShell>();
		LONG lResult;
		VSL_CHECKHRESULT(spIVsUIShell->ShowMessageBox(0, GUID_NULL, strTitle, strMessage, NULL, 0, OLEMSGBUTTON_OK, OLEMSGDEFBUTTON_FIRST, OLEMSGICON_INFO, false, &lResult));

		bHandled = true;
		return 0;
	}

	// Handled to set the color that should be used to draw the background of the Window Pane.
	LRESULT OnCtlColorDlg(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
	{
		if (nullptr != m_hBackground)
		{
			bHandled = true;
		}
		else
		{
			bHandled = false;
		}

		return (LRESULT)m_hBackground;
	}

	STDMETHOD(OnBroadcastMessage)(UINT uMsg, WPARAM /*wParam*/, LPARAM /*lParam*/)
	{
		switch (uMsg)
		{
		case WM_SYSCOLORCHANGE:
		case WM_PALETTECHANGED:
			// Re-initialize VS colors when the theme changes.
			InitVSColors();
			break;
		}

		return S_OK;
	}

private:
	// Initialize colors that are used to render the Window Pane.
	void InitVSColors()
	{
		// Obtain IVsUIShell5 from IVsUIShell
		CComQIPtr<IVsUIShell5> spIVsUIShell5(GetVsSiteCache().GetCachedService<IVsUIShell, SID_SVsUIShell>());
		VS_RGBA vsColor;

		if (nullptr != m_hBackground)
		{
			::DeleteBrush(m_hBackground);
			m_hBackground = nullptr;
		}

		if (nullptr != spIVsUIShell5 && SUCCEEDED(spIVsUIShell5->GetThemedColor(EnvironmentColorsCategory, L"Window", TCT_Background, &vsColor)))
		{
			COLORREF crBackground = VS_RGBA_TO_COLORREF(vsColor);
			m_hBackground = ::CreateSolidBrush(crBackground);
		}
		
		if (::IsWindow(this->m_hWnd))
		{
			::InvalidateRect(this->m_hWnd, nullptr /* lpRect */ , true /* bErase */);
		}
	}

	HBRUSH m_hBackground;
	VSCOOKIE m_BroadcastCookie;
};


class CAppXmlToolWnd :
	public VSL::ToolWindowBase<CAppXmlToolWnd>
{
public:
	IWndPage*			m_pPage;
	IWndFrame*			m_pFrame;
	CAppXmlWndPane*		m_pPane;
	// Constructor of the tool window object.
	// The goal of this constructor is to initialize the base class with the site cache
	// of the owner package.
	CAppXmlToolWnd(const PackageVsSiteCache& rPackageVsSiteCache):
		ToolWindowBase(rPackageVsSiteCache)
	{
		m_pPane = NULL;
		m_pFrame = NULL;
	}

	// Caption of the tool window.
	const wchar_t* const GetCaption() const
	{
		static CStringW strCaption;
		// Avoid to load the string from the resources more that once.
		if (0 == strCaption.GetLength())
		{
			VSL_CHECKBOOL_GLE(
				strCaption.LoadStringW(_AtlBaseModule.GetResourceInstance(), IDS_WINDOW_TITLE));
		}
		return strCaption;
	}

	// Creation flags for this tool window.
	VSCREATETOOLWIN GetCreationFlags() const
	{
		return CTW_fInitNew|CTW_fForceCreate;
	}

	// Return the GUID of the persistence slot for this tool window.
	const GUID& GetToolWindowGuid() const;

	IUnknown* GetViewObject()
	{
		// Should only be called once per-instance
		VSL_CHECKBOOLEAN_EX(m_spView == NULL, E_UNEXPECTED, IDS_E_GETVIEWOBJECT_CALLED_AGAIN);

		// Create the object that implements the window pane for this tool window.
		CComObject<CAppXmlWndPane>* pViewObject;
		VSL_CHECKHRESULT(CComObject<CAppXmlWndPane>::CreateInstance(&pViewObject));
		m_pPane = pViewObject;
		// Get the pointer to IUnknown for the window pane.
		HRESULT hr = pViewObject->QueryInterface(IID_IUnknown, (void**)&m_spView);
		if (FAILED(hr))
		{
			// If QueryInterface failed, then there is something wrong with the object.
			// Delete it and throw an exception for the error.
			delete pViewObject;
			VSL_CHECKHRESULT(hr);
		}

		return m_spView;
	}

	// This method is called by the base class after the tool window is created.
	// We use it to set the icon for this window.
	void PostCreate();

private:
	CComPtr<IUnknown> m_spView;
};
