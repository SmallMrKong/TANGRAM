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
#include <jni.h>
#include "../tangram/CloudAddinCore.h"
class CTangramCtrl;

namespace EclipseCloudPlus
{
	namespace EclipsePlus
	{
		class CEclipseWnd;
		class CEclipseCloudAddin : public CTangram,
			public map<HWND, CEclipseWnd*>
		{
		public:
			CEclipseCloudAddin();
			virtual ~CEclipseCloudAddin();
			bool m_bClose;
			int m_nIndex;
			CString m_strURL;
			void InitEclipsePlus();

			void InitTangramforJava();
		};

		class CEclipseWnd :
			public map<HWND, CTangramCtrl*>,
			public CComObjectRootEx<CComSingleThreadModel>,
			public CWindowImpl<CEclipseWnd, CWindow>,
			public IDispatchImpl<IEclipseTopWnd, &IID_IEclipseTopWnd, &LIBID_Tangram, /*wMajor =*/ 1, /*wMinor =*/ 0>
		{
		public:
			CEclipseWnd(void);
			~CEclipseWnd(void);

			int m_nIndex;
			BOOL m_bCreated;
			HWND m_hClient;
			CString m_strURL;

			IWndPage* m_pPage;
			IWndFrame* m_pFrame;
			IWndNode* m_pCurNode;

			BEGIN_COM_MAP(CEclipseWnd)
				COM_INTERFACE_ENTRY(IEclipseTopWnd)
				COM_INTERFACE_ENTRY(IDispatch)
			END_COM_MAP()

			BEGIN_MSG_MAP(CEclipseWnd)
				MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
				MESSAGE_HANDLER(WM_TANGRAMUCMAMSG, OnUCMAMsg)
			END_MSG_MAP()
			void OnFinalMessage(HWND hWnd);
		protected:
			HWND m_hBottom;
			HWND m_hTopBottom;

			ULONG InternalAddRef() { return 1; }
			ULONG InternalRelease() { return 1; }

			LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
			LRESULT OnUCMAMsg(UINT, WPARAM, LPARAM, BOOL&);
		public:
			STDMETHOD(get_Handle)(long* pVal);
			STDMETHOD(SWTExtend)(BSTR bstrKey, BSTR bstrXml, IWndNode** ppNode);
			STDMETHOD(GetCtrlText)(BSTR bstrNodeName, BSTR bstrCtrlName, BSTR* bstrVal);
		};

		class CEclipseSWTWnd :
			public CWindowImpl<CEclipseSWTWnd, CWindow>
		{
		public:
			CEclipseSWTWnd(void);
			~CEclipseSWTWnd(void);

			CTangramCtrl* m_pHostCtrl;
			BEGIN_MSG_MAP(CEclipseWnd)
				MESSAGE_HANDLER(WM_SWTCOMPONENTNOTIFY, OnSWTComponentNotify)
			END_MSG_MAP()
			void OnFinalMessage(HWND hWnd);
			LRESULT OnSWTComponentNotify(UINT, WPARAM, LPARAM, BOOL&);
		};
	}
}

