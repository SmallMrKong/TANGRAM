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

#ifndef _WIN64
#pragma once
#include "dte.h"
#include "dte80.h"
#include "devshell.h"

#include "../CloudAddinCore.h"
#include "../Tangram/OfficePlus/CloudAddin.h"
#include "../Tangram/VisualStudioPlus/DTEEvent.h"

using namespace VxDTE;
using namespace VSCloudPlus::VisualStudioPlus::VSEvents;

namespace VSCloudPlus
{
	namespace VisualStudioPlus
	{
		class CVSDesignerWnd;
		class CVSComponentWnd :	public CWindowImpl<CVSComponentWnd, CWindow>
		{
		public:
			CVSComponentWnd(void);

			HWND					m_hHostWnd;
			HWND					m_hChildHostWnd;
			BOOL					m_bCloudAddinComponent;
			CString					m_strName;

			IWndPage*				m_pPage;
			IWndNode*				m_pNode;
			CWndFrame*				m_pFrame;
			CVSDesignerWnd*			m_pDesignerWnd;

			BEGIN_MSG_MAP(CVSComponentWnd)
				MESSAGE_HANDLER(WM_TANGRAMDATA, OnGetMe)
				MESSAGE_HANDLER(WM_TANGRAMMSG, OnTangramMsg)
				MESSAGE_HANDLER(WM_NAMECHANGED, OnNameChanged)
				MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
				MESSAGE_HANDLER(WM_TANGRAMGETDESIGNWND, OnGetDesignWnd)
			END_MSG_MAP()
			LRESULT OnGetMe(UINT, WPARAM, LPARAM, BOOL&) { return (LRESULT)this; };
			LRESULT OnGetDesignWnd(UINT, WPARAM, LPARAM, BOOL&);
			LRESULT OnNameChanged(UINT, WPARAM, LPARAM, BOOL&);
			LRESULT OnTangramMsg(UINT, WPARAM, LPARAM, BOOL&);
			LRESULT OnMouseActivate(UINT, WPARAM, LPARAM, BOOL&);
			
			void Init();
			void OnFinalMessage(HWND hWnd);
		};

		class CVSDesignerWnd :
			public CDTEDocumentEvents,
			public map<CString, CVSComponentWnd*>,
			public CWindowImpl<CVSDesignerWnd, CWindow>
		{
		public:
			CVSDesignerWnd(void);

			BOOL							m_bDesignState;
			HWND							m_hMdiClient;
			CString							m_strName;
			CString							m_strPath;

			VxDTE::Document*				m_pDocument;
			DocumentEvents*					m_pDocEvent;
			CComPtr<VxDTE::Project>			m_pPrj;

			IWndPage*						m_pPage;
			IWndNode*						m_pNode;
			CWndFrame*						m_pFrame;

			CVSComponentWnd*				m_pActiveWnd;

			void __stdcall OnDocumentSaved(VxDTE::Document* Document);
			void __stdcall OnDocumentClosing(VxDTE::Document* Document);
			void __stdcall OnDocumentOpening(BSTR DocumentPath, VARIANT_BOOL ReadOnly);
			void __stdcall OnDocumentOpened(VxDTE::Document* Document);

			BEGIN_MSG_MAP(CVSDesignerWnd)
				MESSAGE_HANDLER(WM_TANGRAMDATA, OnGetMe)
				MESSAGE_HANDLER(WM_TANGRAMMSG, OnTangramMsg)
				MESSAGE_HANDLER(WM_NAMECHANGED, OnNameChanged)
				MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
				MESSAGE_HANDLER(WM_TANGRAMGETDESIGNWND, OnGetDesignWnd)
			END_MSG_MAP()
			
			void Init();
			void OnFinalMessage(HWND hWnd);

			LRESULT OnGetMe(UINT, WPARAM, LPARAM, BOOL&) { return (LRESULT)this; };
			LRESULT OnTangramMsg(UINT, WPARAM, LPARAM, BOOL&);
			LRESULT OnNameChanged(UINT, WPARAM, LPARAM, BOOL&);
			LRESULT OnGetDesignWnd(UINT, WPARAM, LPARAM, BOOL&);
			LRESULT OnMouseActivate(UINT, WPARAM, LPARAM, BOOL&);
		};

		class CVSCloudAddin :
			public CTangram,
			public CDTEEvents,
			public CDTEWindowEvents,
			public CDTEProjectsEvents,
			public CDTEProjectItemsEvents,
			public CDTEWindowVisibilityEvents,
			public map<VxDTE::Document*, CVSDesignerWnd*>
		{
		public:
			CVSCloudAddin();
			virtual ~CVSCloudAddin();

			HWND										m_hPropertyForm;
			CString										m_strprjKindVBProject;
			CString										m_strprjKindCSharpProject;
			CVSDesignerWnd*								m_pActiveDesignerWnd;

			_DTE*										m_pDTE;
			//VxDTE::Window*								m_pToolBoxWnd;
			VxDTE::Window*								m_pPropertyWnd;
			VxDTE::Windows*								m_pWindows;
			OutputWindowPane*							m_pOutputWindowPane;

			DTEEvents*									m_pDTEEvents;
			WindowEvents*								m_pWindowEvents;
			ProjectsEvents*								m_pProjectsEvents;
			ProjectItemsEvents*							m_pProjectItemsEvents;
			WindowVisibilityEvents*						m_pPropertyWndVisibilityEvents;

			map<CString, CTangramXmlParse*>				m_mapPrjXml;

			void ClearOutputPane();
			void OnInitInstance();
			void InitPropertyGridCtrl(HWND);
			void AttachDoc(VxDTE::Document*);
			void PutTextToOutputPane(BSTR bstrMsg);
			BOOL IsCloudAddinProject(VxDTE::Project* pPrj);
			BOOL UpdateProjectforCloudAddin(VxDTE::Project* pPrj);
			LRESULT OnTangramDesignMsg(CString );

			//void __stdcall OnStartupComplete();
			void __stdcall OnBeginShutdown();

			//void __stdcall OnWindowClosing(VxDTE::Window* Window);
			void __stdcall OnWindowActivated(VxDTE::Window* GotFocus, VxDTE::Window* LostFocus);
			void __stdcall OnWindowCreated(VxDTE::Window* Window);

			//void __stdcall OnBeforeClosing();
			//void __stdcall OnAfterClosing();

			void __stdcall OnItemAdded(VxDTE::Project* Project);
			void __stdcall OnItemRemoved(VxDTE::Project* Project);
			void __stdcall OnItemRenamed(VxDTE::Project* Project, BSTR OldName);

			void __stdcall OnItemAdded(ProjectItem* ProjectItem);
			void __stdcall OnItemRemoved(ProjectItem* ProjectItem);
			void __stdcall OnItemRenamed(ProjectItem* ProjectItem, BSTR OldName);

			void __stdcall OnWindowHiding(VxDTE::Window* Window);
			void __stdcall OnWindowShowing(VxDTE::Window* Window);
		};

		//class CDSAddin :
		//	public IDSAddIn,
		//	public CTangram
		//{
		//public:
		//	CDSAddin();

		//	BEGIN_COM_MAP(CDSAddin)
		//		COM_INTERFACE_ENTRY2(IDispatch, ITangram)
		//		COM_INTERFACE_ENTRY(ITangram)
		//		COM_INTERFACE_ENTRY(IDSAddIn)
		//		COM_INTERFACE_ENTRY(IConnectionPointContainer)
		//	END_COM_MAP()

		//	DWORD					m_dwCookie;
		//	CComPtr<IApplication>	m_pVSApp;

		//	// IDSAddIn Methods
		//public:
		//	STDMETHOD(OnConnection)(IApplication * pApp, VARIANT_BOOL bFirstTime, long dwCookie, VARIANT_BOOL * OnConnection);
		//	STDMETHOD(OnDisconnection)(VARIANT_BOOL bLastTime);
		//};

		//class CVisualStudioAddin :
		//	public OfficeCloudPlus::CCloudAddinBase,
		//	public CCloudAddinApplication
		//{
		//public:
		//	CVisualStudioAddin();
		//	virtual ~CVisualStudioAddin();

		//	CComPtr<_DTE> m_pDTE;
		//private:
		//	//CCloudAddinApp:
		//	HRESULT Tangram_Command(IDispatch*);
		//	HRESULT OnConnection(IDispatch* pHostApp, int ConnectMode);
		//	HRESULT OnDisconnection(int DisConnectMode);
		//	HRESULT GetCustomAddinUI(BSTR RibbonID, BSTR* bstrXmlScript);
		//};
	}
}
#endif


