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

// WndFrame.h : Declaration of the CTangramNavigator

#pragma once

#include "def.h"       // main symbols
#include "TangramProxy.h"

using namespace System::Reflection;
using namespace TangramCLR;


struct ObjectEventInfo2
{
	map<WindowEventType, IEventProxy*> m_mapEventObj;
};

struct ObjectEventInfo
{
	map<LONGLONG, ObjectEventInfo2*>	m_mapEventObj2;
	map<WindowEventType, IEventProxy*>	m_mapEventObj;
};

class CCloudAddinCLRProxy : public CApplicationCLRProxyImpl
{
public: 
	CCloudAddinCLRProxy();
	 virtual~CCloudAddinCLRProxy();

	 HWND									m_hCurDesignObj;
	 map<HWND, ObjectEventInfo*>			m_mapObjectEventInfo;
	 map<HWND, HWND>						m_mapDesignObj;
	 gcroot<PropertyGrid^>					m_pPropertyGrid;
	 gcroot<PropertyGrid^>					m_pVSPropertyGrid;
	 gcroot<TangramCLR::TangramProxy^>		m_pTangramProxy;

	 Object^ _getObject(Object^ key);
	 bool _insertObject(Object^ key, Object^ val);
	 bool _removeObject(Object^ key);

	 template<typename T1, typename T2>
	 T2^ _createObject(T1* punk)
	 {
		 T2^ pRetObj = nullptr;

		 if (punk != NULL)
		 {
			 LONGLONG l = (LONGLONG)punk;
			 pRetObj = (T2^)_getObject(l);

			 if (pRetObj == nullptr)
			 {
				 pRetObj = gcnew T2(punk);
				 _insertObject(l, pRetObj);
			 }
		 }
		 return pRetObj;
	 }
private:
	 gcroot<Hashtable^> m_htObjects;
	 gcroot<Object^>	m_pTangramObj;
	 gcroot<Assembly^>	m_pSystemAssembly;

	 BSTR AttachObjEvent(IDispatch* EventObj, IDispatch* SourceObj, WindowEventType EventType, IDispatch* HTMLWindow);
	 HRESULT ActiveCLRMethod(BSTR bstrObjID, BSTR bstrMethod, BSTR bstrParam, BSTR bstrData);
	 IDispatch* CreateCLRObj(BSTR bstrObjID);
	 HRESULT ProcessCtrlMsg(HWND hCtrl, bool bShiftKey);
	 BOOL ProcessFormMsg(HWND hFormWnd, LONG lpMsg, int nMouseButton);
	 HRESULT CloseForms();
	 IDispatch* TangramCreateObject(BSTR bstrObjID, long hParent, IWndNode* pHostNode);
	 BOOL IsWinForm(HWND hWnd);
	 IDispatch* GetCLRControl(IDispatch* CtrlDisp, BSTR bstrNames);
	 BSTR GetCtrlName(IDispatch* pCtrl);
	 HWND GetMDIClientHandle(IDispatch* pMDICtrl);
	 IDispatch* GetCtrlByName(IDispatch* CtrlDisp, BSTR bstrName, bool bFindInChild);
	 HWND GetCtrlHandle(IDispatch* pCtrl);
	 HWND IsCtrlCanNavigate(IDispatch* ctrl);
	 void TangramAction(BSTR bstrXml, IWndNode* pNode);
	 BSTR GetCtrlValueByName(IDispatch* CtrlDisp, BSTR bstrName, bool bFindInChild);
	 CString GetCtrlInfo(long hWnd);
	 HWND OnInitialDesigner(long hWnd, CString& strName);
	 Control^ GetCanSelect(Control^ ctrl, bool direct);
	 CString GetControlStrs(CString strLib);
	 CString GetLibControlsStr(String^ strLib);
	 CString GetLibControlsStr(Assembly^ Assembly);
	 void AttachPropertyGridCtrl(HWND hWnd);
	 void SelectNode(IWndNode* );
	 void InitComponent(Control^);
	 static Control^ GetParentByName(Control^ ctrl, String^ name);
	 static void OnControlAdded(System::Object ^sender, System::Windows::Forms::ControlEventArgs ^e);
	 static void OnDockChanged(System::Object ^sender, System::EventArgs ^e);
	 static void OnControlRemoved(System::Object ^sender, System::Windows::Forms::ControlEventArgs ^e);
	 static void OnParentChanged(System::Object ^sender, System::EventArgs ^e);
	 static void OnSelectedGridItemChanged(System::Object ^sender, System::Windows::Forms::SelectedGridItemChangedEventArgs ^e);
	 static void OnPropertyValueChanged(System::Object ^s, System::Windows::Forms::PropertyValueChangedEventArgs ^e);
};

extern _ATL_FUNC_INFO Initialize;
extern _ATL_FUNC_INFO Destroy;

class CWndPageClrEvent : public IDispEventSimpleImpl</*nID =*/ 1, CWndPageClrEvent, &__uuidof(_IWndPage)>
{
public:
	CWndPageClrEvent();
	virtual ~CWndPageClrEvent();

	gcroot<Object^> m_pPage;

	void __stdcall  OnDestroy();
	void __stdcall  OnInitialize(IDispatch* pHtmlWnd, BSTR bstrUrl);
	BEGIN_SINK_MAP(CWndPageClrEvent)
		SINK_ENTRY_INFO(/*nID =*/ 1, __uuidof(_IWndPage), /*dispid =*/ 0x00000001, OnInitialize, &Initialize)
		SINK_ENTRY_INFO(/*nID =*/ 1, __uuidof(_IWndPage), /*dispid =*/ 0x00000006, OnDestroy, &Destroy)
	END_SINK_MAP()
};
