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
* OUTLINED IN THE TANGRAM LICENSE AGREEMENT.TANGRAM TEAM
* GRANTS TO YOU (ONE SOFTWARE DEVELOPER) THE LIMITED RIGHT TO USE
* THIS SOFTWARE ON A SINGLE COMPUTER.
*
* CONTACT INFORMATION:
* mailto:sunhuizlz@yeah.net
* http://www.CloudAddin.com
*
********************************************************************************/

// TangramAddin.h : Declaration of the CCloudAddin
#include "..\CloudAddinCore.h"
#pragma once

typedef map<LONG, Office::_CustomTaskPane*> CTaskPaneMap;
typedef CTaskPaneMap::iterator TaskPaneMapIT;

namespace OfficeCloudPlus
{
	class COfficeObject;

	class CTangramDocument
	{
	public:
		CTangramDocument()
		{
			m_strDocXml = _T("");
			m_strTaskPaneXml = _T("");
			m_strTaskPaneTitle = _T("");

			m_pPage = nullptr;
			m_pFrame = nullptr;
			m_pTaskPaneFrame = nullptr;
			m_pTaskPanePage = nullptr;
			m_nMsoCTPDockPosition = msoCTPDockPositionRight;
			m_nMsoCTPDockPositionRestrict = msoCTPDockPositionRestrictNone;
		}

		int								m_nWidth;
		int								m_nHeight;
		MsoCTPDockPosition				m_nMsoCTPDockPosition;
		MsoCTPDockPositionRestrict		m_nMsoCTPDockPositionRestrict;
		CString							m_strDocXml;
		CString							m_strTaskPaneXml;
		CString							m_strTaskPaneTitle;

		IWndPage*						m_pPage;
		IWndPage*						m_pTaskPanePage;
		CWndFrame*						m_pFrame;
		CWndFrame*						m_pTaskPaneFrame;
		map<CString, CString>			m_mapUserFormScript;
	};

	// CCloudAddin
	class ATL_NO_VTABLE CCloudAddin :
		public CTangram
	{
	public:
		CCloudAddin();
		virtual ~CCloudAddin();

		CString							m_strUser;
		CString							m_strRibbonXml;
		CString							m_strTemplateXml;
		CComPtr<Office::ICTPFactory>	m_pCTPFactory;
		CTaskPaneMap					m_mapTaskPaneMap;
		map<HWND, CWndFrame*>			m_mapVBAForm;
		map<HWND, COfficeObject*>		m_mapOfficeObjects;
		map<HWND, COfficeObject*>		m_mapOfficeObjects2;

		void OnCloseOfficeObj(CString strName, HWND hWnd);
		void _AddDocXml(Office::_CustomXMLParts* pCustomXMLParts, BSTR bstrXml, BSTR bstrKey, VARIANT_BOOL* bSuccess);
		BEGIN_COM_MAP(CCloudAddin)
			COM_INTERFACE_ENTRY(ITangram)
			COM_INTERFACE_ENTRY2(IDispatch, ITangram)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		virtual void SetFocus(HWND) {};
		virtual void ConnectOfficeObj(HWND hWnd) {};
		virtual void OnVbaFormInit(HWND, CWndFrame*) {};
		virtual void WindowCreated(LPCTSTR strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd) {};
		virtual void WindowDestroy(HWND hWnd) {};
		virtual bool OnActiveOfficeObj(HWND hWnd) { return false; };
		virtual HRESULT Tangram_OnLoad(IDispatch* RibbonControl) { return S_OK; };
		virtual HRESULT Tangram_GetItemCount(IDispatch* RibbonControl, long* nCount) { return S_OK; };
		virtual HRESULT Tangram_GetItemLabel(IDispatch* RibbonControl, long nIndex, BSTR* bstrLabel) { return S_OK; };
		virtual HRESULT Tangram_GetItemID(IDispatch* RibbonControl, long nIndex, BSTR* bstrID) { return S_OK; };
		virtual HRESULT OnConnection(IDispatch* pHostApp, int ConnectMode) {};
		virtual HRESULT OnDisconnection(int DisConnectMode) {};
		virtual HRESULT OnUpdate(void) { return S_OK; };
		virtual HRESULT StartupComplete(void) { return S_OK; };
		virtual HRESULT BeginShutdown(void) { return S_OK; };
		virtual CString GetFormXml(CString strFormName) { return _T(""); };
	protected:
		CString					m_strLib;
		IWndPage*				m_pPage;
		CComQIPtr<IRibbonUI>	m_spRibbonUI;
		void _GetDocXmlByKey(Office::_CustomXMLParts* pCustomXMLParts, BSTR bstrKey, BSTR* pbstrXml);
	private:

		STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);
		//IDTExtensibility2 implementation:
		STDMETHOD(OnConnection)(IDispatch * Application, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
		STDMETHOD(OnDisconnection)(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY **custom);
		STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom);
		STDMETHOD(OnStartupComplete)(SAFEARRAY **custom);
		STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom);
		//IRibbonExtensibility implementation

		STDMETHOD(CTPFactoryAvailable)(ICTPFactory * CTPFactoryInst);

		STDMETHOD(InitVBAForm)(IDispatch*, long, BSTR bstrXml, IWndNode** ppRetNode);
		STDMETHOD(AddVBAFormsScript)(IDispatch* OfficeObject, BSTR bstrKey, BSTR bstrXml);
		STDMETHOD(ExportOfficeObjXml)(IDispatch* OfficeObject, BSTR* bstrXml);
		STDMETHOD(GetFrameFromVBAForm)(IDispatch* pForm, IWndFrame** ppFrame);
		STDMETHOD(GetActiveTopWndNode)(IDispatch* pForm, IWndNode** WndNode);
		STDMETHOD(GetObjectFromWnd)(LONG hWnd, IDispatch** ppObjFromWnd);
		STDMETHOD(TangramCommand)(IDispatch* RibbonControl);
		STDMETHOD(TangramGetImage)(BSTR strValue, IPictureDisp ** ppDispImage);
		STDMETHOD(TangramGetVisible)(IDispatch* RibbonControl, VARIANT* varVisible);
		STDMETHOD(TangramOnLoad)(IDispatch* RibbonControl);
		STDMETHOD(TangramGetItemCount)(IDispatch* RibbonControl, long* nCount);
		STDMETHOD(TangramGetItemLabel)(IDispatch* RibbonControl, long nIndex, BSTR* bstrLabel);
		STDMETHOD(TangramGetItemID)(IDispatch* RibbonControl, long nIndex, BSTR* bstrID);
		STDMETHOD(get_RemoteHelperHWND)(long* pVal);
	};

	class COfficeObject :
		public CComObjectRootBase,
		public IConnectionPointContainerImpl<COfficeObject>,
		public IDispatchImpl<IOfficeObject, &IID_IOfficeObject, &LIBID_Tangram, 1, 0>
	{
	public:
		COfficeObject(void);
		virtual ~COfficeObject(void);

		BOOL				m_bMDIClient;
		HWND				m_hForm;
		HWND				m_hClient;
		HWND				m_hChildClient;
		HWND				m_hTaskPane;
		HWND				m_hTaskPaneWnd;
		HWND				m_hTaskPaneChildWnd;
		IWndPage*			m_pPage;
		IWndFrame*			m_pFrame;
		IDispatch*			m_pOfficeObj;

		BEGIN_COM_MAP(COfficeObject)
			COM_INTERFACE_ENTRY(IOfficeObject)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(COfficeObject)
		END_CONNECTION_POINT_MAP()

		STDMETHOD(Show)(VARIANT_BOOL bShow);
		STDMETHOD(Extend)(BSTR bstrKey, BSTR bstrXml, IWndNode** ppNode);
		STDMETHOD(UnloadTangram)();

		void Lock() {}
		void Unlock() {}
		virtual void OnObjDestory() {};
	protected:
		ULONG InternalAddRef() { return 1; }
		ULONG InternalRelease() { return 1; }
	};
}

