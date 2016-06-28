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


// WndNode.h : Declaration of the CWndNode

#include "browser.h"
#include "WndPage.h"

#pragma once
class CNodeWnd;
class CWndNodeEvents;
class CTangramHtmlTreeWnd;
class CWndNodeWBEvent : 
	public CJSExtender,
	public CWebbrowserEvent
{
public:
	CWndNodeWBEvent(CWndNode* pNode);
	IHTMLDocument2*	m_pHTMLDocument2;

private:
	CWndNode* m_pWndNode;
	void __stdcall OnDocumentComplete(IDispatch *pDisp, VARIANT *URL);
	void __stdcall OnNavigateComplete2(IDispatch* pDisp, VARIANT* URL);
};

class ATL_NO_VTABLE CEventProxy : 
	public CComObjectRootBase,
	public IConnectionPointContainerImpl < CEventProxy >,
	public IConnectionPointImpl<CEventProxy, &DIID__IEventProxy>,
	public IDispatchImpl<IEventProxy, &IID_IEventProxy, &LIBID_Tangram,  1,  0>
{
public:
	BEGIN_COM_MAP(CEventProxy)
		COM_INTERFACE_ENTRY(IEventProxy)
		COM_INTERFACE_ENTRY(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(CEventProxy)
		CONNECTION_POINT_ENTRY(DIID__IEventProxy)
	END_CONNECTION_POINT_MAP()

	void Lock(){}
	void Unlock(){}

protected:
	ULONG InternalAddRef(){return 1;}
	ULONG InternalRelease(){return 1;}
};

// CWndNode 
class ATL_NO_VTABLE CWndNode : 
	public CComObjectRootBase,
	public IConnectionPointContainerImpl<CWndNode>,
	public IConnectionPointImpl<CWndNode, &__uuidof(_IWndNodeEvents)>,
	public IDispatchImpl<IWndNode, &IID_IWndNode, &LIBID_Tangram,  1,  0>
{
public:
	CWndNode();
	virtual ~CWndNode();

	BOOL								m_bTopObj;
	BOOL								m_bCreated;
	BOOL								m_bWebInit;
	BOOL								m_bNodeDocComplete;

	ViewType							m_nViewType;
	int									m_nID;
	int									m_nActivePage;
	int									m_nWidth;
	int									m_nHeigh;
	int									m_nRow;
	int									m_nCol;
	int									m_nRows;
	int									m_nCols;
	long								m_nStyle;
	HWND								m_hHostWnd;
	HWND								m_hChildHostWnd;

	CString 							m_strID;
	CString 							m_strURL;
	CString								m_strKey;
	CString								m_strName;
	CString 							m_strCnnID;
	CString 							m_strCaption;

	CString								m_strWebObjName;
	CString 							m_strExtenderID;

	IDispatch*							m_pOfficeObj;
	IDispatch*							m_pDisp;
	CWndNode* 							m_pRootObj;
	CWndNode* 							m_pParentObj;
	CWndNode*							m_pVisibleXMLObj;
	CTangramXmlParse* 					m_pDocXmlParseNode;
	CTangramHtmlTreeWnd* 				m_pDocTreeCtrl;

	CTangramXmlParse*					m_pHostParse;
	CTangramXmlParse*					m_pTangramParse;

	CWnd*								m_pHostWnd;
	CWndPage*							m_pPage;
	CNodeWnd*							m_pHostClientView;
	CWndFrame*							m_pFrame;
	IWndFrame*							m_pHostFrame;
	CRuntimeClass*						m_pObjClsInfo;

	CWndNodeEvents*						m_pCLREventConnector;
	CWndNodeWBEvent*					m_pTangramNodeWBEvent;

	CMapStringToPtr						m_PlugInDispDictionary;

	VARIANT								m_varTag;
	IDispatch*							m_pExtender;
	CTangramNodeVector					m_vChildNodes;
	CWndNode*							m_pCurrentExNode;
	map<CString, CWndNode*>				m_mapLayoutNodes;
	CComObject<CWndNodeCollection>*		m_pChildNodeCollection;

	BOOL	PreTranslateMessage(MSG* pMsg);
	BOOL	AddChildNode(CWndNode* pNode);
	BOOL	RemoveChildNode(CWndNode* pNode);

	HRESULT Fire_ExtendComplete()
	{
		HRESULT hr = S_OK;
		int cConnections = m_vec.GetSize();

		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant varResult;

				DISPPARAMS params = { NULL, NULL, 0, 0 };
				hr = pConnection->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
		return hr;
	}

	HRESULT Fire_Destroy()
	{
		HRESULT hr = S_OK;
		int cConnections = m_vec.GetSize();

		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant varResult;

				DISPPARAMS params = { NULL, NULL, 0, 0 };
				hr = pConnection->Invoke(2, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
		return hr;
	}

	HRESULT Fire_NodeAddInCreated(IDispatch * pAddIndisp, BSTR bstrAddInID, BSTR bstrAddInXml)
	{
		HRESULT hr = S_OK;
		int cConnections = m_vec.GetSize();

		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant avarParams[3];
				avarParams[2] = pAddIndisp;
				avarParams[2].vt = VT_DISPATCH;
				avarParams[1] = bstrAddInID;
				avarParams[1].vt = VT_BSTR;
				avarParams[0] = bstrAddInXml;
				avarParams[0].vt = VT_BSTR;
				CComVariant varResult;

				DISPPARAMS params = { avarParams, NULL, 3, 0 };
				hr = pConnection->Invoke(3, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
		return hr;
	}

	HRESULT Fire_NodeAddInsCreated()
	{
		HRESULT hr = S_OK;
		int cConnections = m_vec.GetSize();

		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant varResult;

				DISPPARAMS params = { NULL, NULL, 0, 0 };
				hr = pConnection->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
		return hr;
	}

	HRESULT Fire_NodeDocumentComplete(IDispatch * ExtenderDisp, BSTR bstrURL)
	{
		HRESULT hr = S_OK;
		int cConnections = m_vec.GetSize();

		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant avarParams[2];
				avarParams[1] = ExtenderDisp;
				avarParams[1].vt = VT_DISPATCH;
				avarParams[0] = bstrURL;
				avarParams[0].vt = VT_BSTR;
				CComVariant varResult;

				DISPPARAMS params = { avarParams, NULL, 2, 0 };
				hr = pConnection->Invoke(5, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
		return hr;
	}

	HRESULT Fire_ControlNotify(IWndNode * sender, LONG NotifyCode, LONG CtrlID, LONGLONG CtrlHandle, BSTR CtrlClassName)
	{
		HRESULT hr = S_OK;
		int cConnections = m_vec.GetSize();

		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant avarParams[5];
				avarParams[4] = sender;
				avarParams[3] = NotifyCode;
				avarParams[3].vt = VT_I4;
				avarParams[2] = CtrlID;
				avarParams[2].vt = VT_I4;
				avarParams[1] = CtrlHandle;
				avarParams[1].vt = VT_I8;
				avarParams[0] = CtrlClassName;
				avarParams[0].vt = VT_BSTR;
				CComVariant varResult;

				DISPPARAMS params = { avarParams, NULL, 5, 0 };
				hr = pConnection->Invoke(6, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
		return hr;
	}

	void Lock(){}
	void Unlock(){}
protected:
	ULONG InternalAddRef(){ return 1; }
	ULONG InternalRelease(){ return 1; }

public:
	STDMETHOD(get_XObject)(VARIANT* pVar);
	STDMETHOD(get_Tag)(VARIANT* pVar);
	STDMETHOD(put_Tag)(VARIANT var);
	STDMETHOD(get_AxPlugIn)(BSTR bstrName, IDispatch** pVal);
	STDMETHOD(get_Name)(BSTR* pVal);
	STDMETHOD(put_Name)(BSTR bstrName);
	STDMETHOD(get_Caption)(BSTR* pVal);
	STDMETHOD(put_Caption)(BSTR bstrCaption);
	STDMETHOD(get_Attribute)(BSTR bstrKey, BSTR* pVal);
	STDMETHOD(put_Attribute)(BSTR bstrKey, BSTR bstrVal);
	STDMETHOD(get_Handle)(long* hWnd);
	STDMETHOD(get_XML)(BSTR* pVal);
	STDMETHOD(get_ChildNodes)(IWndNodeCollection** );
	STDMETHOD(get_Rows)(long* nRows);
	STDMETHOD(get_Cols)(long* nCols);
	STDMETHOD(get_Row)(long* nRow);
	STDMETHOD(get_Col)(long* nCol);
	STDMETHOD(get_Objects)(long nType, IWndNodeCollection** );
	STDMETHOD(get_NodeType)(WndNodeType* nType);
	STDMETHOD(get_ParentNode)(IWndNode** ppNode);
	STDMETHOD(get_RootNode)(IWndNode** ppNode);
	STDMETHOD(get_OuterXml)(BSTR* pVal);
	STDMETHOD(get_Key)(BSTR* pVal);
	STDMETHOD(get_Frame)(IWndFrame** pVal);
	STDMETHOD(get_Height)(LONG* pVal);
	STDMETHOD(get_Width)(LONG* pVal);
	STDMETHOD(get_Extender)(IDispatch** pVal);
	STDMETHOD(put_Extender)(IDispatch* newVal);
	STDMETHOD(get_WndPage)(IWndPage** pVal);
	STDMETHOD(get_HTMLWindow)(IDispatch** pVal);
	STDMETHOD(get_HtmlDocument)(IDispatch** pVal);
	STDMETHOD(get_NameAtWindowPage)(BSTR* pVal);
	STDMETHOD(get_Hmin)(int* pVal);
	STDMETHOD(put_Hmin)(int newVal);
	STDMETHOD(get_Hmax)(int* pVal);
	STDMETHOD(put_Hmax)(int newVal);
	STDMETHOD(get_Vmin)(int* pVal);
	STDMETHOD(put_Vmin)(int newVal);
	STDMETHOD(get_Vmax)(int* pVal);
	STDMETHOD(put_Vmax)(int newVal);
	STDMETHOD(get_DocXml)(BSTR* pVal);
	STDMETHOD(get_rgbMiddle)(OLE_COLOR* pVal);
	STDMETHOD(put_rgbMiddle)(OLE_COLOR newVal);
	STDMETHOD(get_rgbRightBottom)(OLE_COLOR* pVal);
	STDMETHOD(put_rgbRightBottom)(OLE_COLOR newVal);
	STDMETHOD(get_rgbLeftTop)(OLE_COLOR* pVal);
	STDMETHOD(put_rgbLeftTop)(OLE_COLOR newVal);

	STDMETHOD(ActiveTabPage)(IWndNode* pNode);
	STDMETHOD(Extend)(BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode);
	STDMETHOD(ExtendEx)(int nRow, int nCol, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode);
	STDMETHOD(GetNode)(long nRow, long nCol,IWndNode** ppTangramNode);
	STDMETHOD(GetNodes)(BSTR bstrName, IWndNode** ppNode, IWndNodeCollection** ppNodes, long* pCount);
	STDMETHOD(GetCtrlByName)(BSTR bstrName, VARIANT_BOOL bFindInChild, IDispatch** ppRetDisp);
	STDMETHOD(Refresh)(void);
	STDMETHOD(LoadXML)(int nType, BSTR bstrXML);
	STDMETHOD(Show)();
	STDMETHOD(UpdateDesignerData)(BSTR bstrXML);

	BEGIN_COM_MAP(CWndNode)
		COM_INTERFACE_ENTRY(IWndNode)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(CWndNode)
		CONNECTION_POINT_ENTRY(__uuidof(_IWndNodeEvents))
	END_CONNECTION_POINT_MAP()

	HWND CreateView(HWND hParentWnd, CString strTag);

private:
	void _get_Objects(CWndNode* pNode, UINT32& nType, CWndNodeCollection* pNodeColletion);
	int _getNodes(CWndNode* pNode, CString& strName, CWndNode**ppRetNode, CWndNodeCollection* pNodes);
};

// CWndNodeCollection

class ATL_NO_VTABLE CWndNodeCollection :
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDispatchImpl<IWndNodeCollection, &IID_IWndNodeCollection, &LIBID_Tangram, 1, 0>
{
public:
	CWndNodeCollection();
	~CWndNodeCollection();

	BEGIN_COM_MAP(CWndNodeCollection)
		COM_INTERFACE_ENTRY(IWndNodeCollection)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()

	STDMETHOD(get_NodeCount)(long* pCount);
	STDMETHOD(get_Item)(long iIndex, IWndNode **ppTangramNode);
	STDMETHOD(get__NewEnum)(IUnknown** ppVal);
	CTangramNodeVector*	m_pNodes;

private:
	CTangramNodeVector	m_vNodes;
};
