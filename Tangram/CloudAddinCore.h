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

#include "tangram.h"
#include "ObjSafe.h"

// CloudAddinCore.h : Declaration of the CTangram

#pragma once

class CWebTaskObj;
class CTangramHtmlTreeWnd;

// CTangram

class ATL_NO_VTABLE CTangram : 
	public CComObjectRootBase,
	public IConnectionPointContainerImpl<CTangram>,
	public IConnectionPointImpl<CTangram, &__uuidof(_ITangramEvents)>,
	public IDispatchImpl<ITangram, &IID_ITangram, &LIBID_Tangram,  1,  0>
{
public:
	CTangram();
	virtual ~CTangram();

	bool									m_bCanClose;
	int										m_nRef;
	HWND									m_hMainWnd;
	HWND									m_hHostWnd;
	HWND									m_hChildHostWnd;
	HWND									m_hMainFrameWnd;

	map<CString, CString>					m_strMapKey;
	map<CString, IDispatch*>				m_mapObjDic;
	map<CString, IDispatch*>				m_mapAppDispDic;
	map<CString, CWebTaskObj*>				m_mapWebTask;

	BEGIN_COM_MAP(CTangram)
		COM_INTERFACE_ENTRY(ITangram)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(CTangram)
		CONNECTION_POINT_ENTRY(__uuidof(_ITangramEvents))
	END_CONNECTION_POINT_MAP()

	BEGIN_CATEGORY_MAP(CTangram)
		IMPLEMENTED_CATEGORY(CATID_SafeForInitializing)
		IMPLEMENTED_CATEGORY(CATID_SafeForScripting)
	END_CATEGORY_MAP()

	STDMETHOD(put_ExternalInfo)(VARIANT newVal);
	STDMETHOD(get_RootNodes)(IWndNodeCollection** pNodeColletion);
	STDMETHOD(get_CurrentActiveNode)(IWndNode** pVal);
	STDMETHOD(get_CreatingNode)(IWndNode** pVal);
	STDMETHOD(get_Application)(IDispatch** pVal);
	STDMETHOD(get_AppExtender)(BSTR bstrKey, IDispatch** pVal);
	STDMETHOD(put_AppExtender)(BSTR bstrKey, IDispatch* newVal);
	STDMETHOD(get_AppKeyValue)(BSTR bstrKey, VARIANT* pVal);
	STDMETHOD(put_AppKeyValue)(BSTR bstrKey, VARIANT newVal);
	STDMETHOD(get_RemoteHelperHWND)(long* pVal) { *pVal = (long)m_hHostWnd; return S_OK; };
	STDMETHOD(put_RemoteHelperHWND)(long newVal) { m_hHostWnd = (HWND)newVal; return S_OK; };
	STDMETHOD(get_DesignNode)(IWndNode** pVal);
	STDMETHOD(put_CurrentDesignNode)(IWndNode* newVal);
	STDMETHOD(CreateCLRObj)(BSTR bstrObjID, IDispatch** ppDisp);
	STDMETHOD(GetWndFrame)(long hHostWnd, IWndFrame** ppFrame);
	STDMETHOD(SetHostFocus)(void);
	STDMETHOD(Encode)(BSTR bstrSRC, VARIANT_BOOL bEncode, BSTR* bstrRet);
	STDMETHOD(NewGUID)(BSTR* retVal);
	STDMETHOD(ActiveCLRMethod)(BSTR bstrObjID, BSTR bstrMethod, BSTR bstrParam, BSTR bstrData);
	STDMETHOD(GetCLRControl)(IDispatch* CtrlDisp, BSTR bstrNames, IDispatch** ppRetDisp);
	STDMETHOD(GetItemText)(IWndNode* pNode, long nCtrlID, LONG nMaxLengeh, BSTR* bstrRet);
	STDMETHOD(SetItemText)(IWndNode* pNode, long nCtrlID, BSTR bstrRet);
	STDMETHOD(CreateWndPage)(long hWnd, IWndPage** ppTangram);
	STDMETHOD(CreateWebTask)(BSTR bstrTaskName, BSTR bstrTaskURL, IWebTaskObj** pWebTaskObj);
	STDMETHOD(CreateTaskObj)(ITaskObj** ppTaskObj);
	STDMETHOD(CreateRestObj)(IRestObject** ppRestObj);
	STDMETHOD(TangramGetObject)(IDispatch* SourceDisp, IDispatch** ResultDisp);
	STDMETHOD(ConnectApp)(BSTR bstrAppID, ITangram** ResultTangramCore);
	STDMETHOD(CreateOfficeObj)(BSTR bstrAppID, BSTR bstrXml);
	STDMETHOD(MessageBox)(long hWnd, BSTR bstrContext, BSTR bstrCaption, long nStyle, int* nRet);
	STDMETHOD(DownLoadFile)(BSTR strFileURL, BSTR bstrTargetFile, BSTR bstrActionXml);
	STDMETHOD(ExtendXml)(BSTR bstrXml, BSTR bstrKey, IDispatch** ppNode);
	STDMETHOD(TangramCommand)(IDispatch* RibbonControl) { return S_OK; };
	STDMETHOD(TangramGetImage)(BSTR strValue, IPictureDisp ** ppDispImage) { return S_OK; };
	STDMETHOD(TangramGetVisible)(IDispatch* RibbonControl, VARIANT* varVisible) { return S_OK; };
	STDMETHOD(TangramOnLoad)(IDispatch* RibbonControl) { return S_OK; };
	STDMETHOD(TangramGetItemCount)(IDispatch* RibbonControl, long* nCount) { return S_OK; };
	STDMETHOD(TangramGetItemLabel)(IDispatch* RibbonControl, long nIndex, BSTR* bstrLabel) { return S_OK; };
	STDMETHOD(TangramGetItemID)(IDispatch* RibbonControl, long nIndex, BSTR* bstrID) { return S_OK; };
	STDMETHOD(AddDocXml)(IDispatch* pDocdisp, BSTR bstrXml, BSTR bstrKey, VARIANT_BOOL* bSuccess) { return S_OK; };
	STDMETHOD(GetDocXmlByKey)(IDispatch* pDocdisp, BSTR bstrKey, BSTR* pbstrXml) { return S_OK; };
	STDMETHOD(CreateOfficeDocument)(BSTR bstrXml) { return S_OK;};
	STDMETHOD(UpdateWndNode)(IWndNode* pNode);
	STDMETHOD(GetNodeFromeHandle)(long hWnd, IWndNode** ppRetNode);
	STDMETHOD(GetCLRControlString)(BSTR bstrAssemblyPath, BSTR* bstrCtrls);
	STDMETHOD(GetNewLayoutNodeName)(BSTR strCnnID, IWndNode* pDesignNode, BSTR* bstrNewName);
	STDMETHOD(InitVBAForm)(IDispatch*, long, BSTR bstrXml, IWndNode** ppRetNode);
	STDMETHOD(AddVBAFormsScript)(IDispatch* OfficeObject, BSTR bstrKey, BSTR bstrXml);
	STDMETHOD(ExportOfficeObjXml)(IDispatch* OfficeObject, BSTR* bstrXml);
	STDMETHOD(GetFrameFromVBAForm)(IDispatch* pForm, IWndFrame** ppFrame);
	STDMETHOD(GetActiveTopWndNode)(IDispatch* pForm, IWndNode** WndNode);
	STDMETHOD(GetObjectFromWnd)(LONG hWnd, IDispatch** ppObjFromWnd);
	STDMETHOD(ReleaseTangram)();
	STDMETHOD(InitJava)(int nIndex);

	void Init();
	void Lock(){}
	void Unlock(){}
	void CreateCommonDesignerToolBar();

	virtual void ProcessMsg(LPMSG lpMsg) {};
	virtual void OnOpenDoc(WPARAM) {};
	virtual void UpdateOfficeObj(IDispatch* pObj, CString strXml, CString strName) {};
	CWndNode* ExtendEx(long hHostMainWnd, CString strXTMLFile, LONG l, LONG t, LONG r, LONG b, LONG l2, LONG t2, LONG r2, LONG b2);

	CString GetDocTemplateXml(CString strCaption, CString strPath);

protected:
	CString	m_strDefaultTemplateXml;
	ULONG InternalAddRef(){ return 1; }
	ULONG InternalRelease();
private:
	IWndPage*							m_pPage;
	IWndFrame*							m_pFrame;
	CWindow								m_HelperWnd;
	CComObject<CWndNodeCollection>*		m_pRootNodes;
public:
	STDMETHOD(put_JavaProxy)(IJavaProxy* newVal);
	STDMETHOD(get_HostWnd)(LONG* pVal);
	STDMETHOD(put_CollaborationProxy)(ICollaborationProxy* newVal);
};
