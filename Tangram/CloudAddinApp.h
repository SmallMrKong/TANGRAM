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


// CloudAddinApp.h : DLL

#pragma once
#include <atlimage.h>
#include "metahost.h"
#include <jni.h>
class CRestObj;
class CHelperWnd;

typedef HRESULT (__stdcall *TangramCLRCreateInstance)(REFCLSID clsid, REFIID riid, /*iid_is(riid)*/ LPVOID *ppInterface);
struct TangramThreadInfo
{
	HHOOK						m_hGetMessageHook;
	map<HWND, CWndFrame*>		m_mapTangramFrame;
};

struct TangramWndClsInfo
{
public:
	TangramWndClsInfo(void){ m_nType = TangramView; };

	LPCTSTR			m_strObjID;
	LPCTSTR			m_strCnnID;
	ViewType		m_nType;
	CRuntimeClass*	m_pTabWndClsInfo;
};

class CTangramHtmlTreeWnd;

class CHelperWnd :
	public CWindowImpl<CHelperWnd, CWindow>
{
public:
	CHelperWnd(void) { m_strID = _T(""); };
	~CHelperWnd(void) {};

	CString m_strID;
	DECLARE_WND_CLASS(_T("Tangram Helper Window Class"))
	BEGIN_MSG_MAP(CHelperWnd)
	END_MSG_MAP()
	void OnFinalMessage(HWND hWnd);
};

class CTangram;
class CTangramApp:
	public CWinApp,
	public CTangramProxyBase,
	public CAtlDllModuleT< CTangramApp >
{
public:
	CTangramApp();
	~CTangramApp();

	BOOL									m_bClose;
	BOOL									m_bMSAddin;
	BOOL									m_bOfficeApp;
	BOOL									m_bCLRStart;
	BOOL									m_b32Process;
	BOOL									m_bWinFormActived;
	BOOL									m_bInitTangramJava;
	BOOL									m_bEnableProcessFormTabKey;

	int										m_nAppID;
	DWORD									m_dwThreadID;
#ifdef _DEBUG
	int										m_nTangram;
	int										m_nTangramObj;
	int										m_nTangramFrame;
	int										m_nOfficeDocs;
	int										m_nOfficeDocsSheet;
#endif
	IDispatch*								m_pAppDisp;
	ITangram*								m_pTangram;
	ITangramAddin*							m_pTangramAddin;
	CWndPage*								m_pPage;
	CWndNode*								m_pDesignWindowNode;
	CWndNode*								m_pHostDesignUINode;
	CWndNode*								m_pWndNode;
	CWndFrame*								m_pWndFrame;
	CWndFrame*								m_pDesigningFrame;
	CTangram*								m_pHostCore;
	CTangramProxy*							m_pTangramProxy;
	IJavaProxy*								m_pJavaProxy;
	ICollaborationProxy*					m_pICollaborationProxy;

	CEventProxy*							m_pEventProxy; 
	CBKWnd*									m_pMDIClientBKWnd;
	CComPtr<ITypeInfo>						m_pEventTypeInfo;
	CTangramPackageProxy*					m_pTangramPackageProxy;
	CApplicationCLRProxyImpl*				m_pCloudAddinCLRProxy;

	CTangramXmlParse*						m_pCurWorkNodeParse;

	CWndFrame*								m_pDesignerFrame;
	IWndPage*								m_pDesignerWndPage;
	CTangramHtmlTreeWnd*					m_pDocDOMTree;

	JNIEnv *								m_pJVMenv;
	JavaVM * 								m_pJVM;

	HWND									m_hActiveWnd;

	HHOOK									m_hCBTHook;
	HMODULE									m_hModule;
	//BOOL m_bCEFInitialized;
	//CefRefPtr<ClientApp> m_cefApp;

	//CString m_szHomePage;

	TCHAR									m_szBuffer[MAX_PATH];
	LPCTSTR									m_lpszSplitterClass;

	CString									m_strExeName;
	CString									m_strAppPath;
	CString									m_strTempPath;
	CString 								m_strConfigFile;
	CString									m_strCurrentKey;
	CString									m_strAppDataPath;
	CString									m_strTangramCLRPath;
	CString									m_strProgramFilePath;
	CString									m_strStartJar;
	VARIANT									m_varApplication;

	map<HWND, CWndPage*>					m_mapWindowPage;
	map<CString, int>						m_mapOfficeAppID;
	map<CString, int>						m_mapOfficeExeID;
	map<CString, CComVariant>				m_mapValInfo;
	map<CString, ITangram*>					m_mapRemoteTangramCore;
	map<CString, WindowEventType>			m_mapEventType;
	map<CString, CHelperWnd*>				m_mapRemoteTangramHelperWnd;
	vector<CRestObj*>						m_vCloudAddinRestObj;
	
	void InitJava(int nIndex);
	int	 LoadCLR();
	LRESULT Close(void);
	void InitEventDic();
	CString GetNewGUID();
	CString GetFileVer();
	BOOL IsUserAdministrator();
	void SetHook(DWORD ThreadID);
	bool CheckUrl(CString&   url);
	CString GetNames(IWndNode* pNode);
	void InitWndNode(CWndNode* pObj);
	void InitDesignerTreeCtrl(CString strXml);
	CString EncodeFileToBase64(CString strSRC);
	CString Encode(CString strSRC, BOOL bEnCode);
	void HostPosChanged(CWndNode* pHostViewNode);
	TangramThreadInfo* GetThreadInfo(DWORD dwInfo=0);
	CString GetXmlData(CString strName, CString strXml);
	DWORD ExecCmd(const CString cmd, const BOOL setCurrentDirectory);
	CString GetPropertyFromObject(IDispatch* pObj, CString strPropertyName);
	BOOL LoadImageFromResource(ATL::CImage *pImage, HMODULE hMod, UINT nResID, LPCTSTR lpTyp);
	BOOL Create(CWndNode* pWindowNode, LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext = NULL);

	static LRESULT CALLBACK CBTProc(int nCode, WPARAM wParam, LPARAM lParam);
	static LRESULT CALLBACK TangramWndProc(_In_ HWND hWnd, UINT msg,_In_ WPARAM wParam,_In_ LPARAM lParam);

	void _addObject(void* pThis, IUnknown* pUnknown)
	{
		m_mapObjects.insert(make_pair(pThis,pUnknown));
	}

	void _removeObject(void* pThis)
	{
		auto it = m_mapObjects.find(pThis);
		if (it != m_mapObjects.end())
		{
			m_mapObjects.erase(it);
		}
	}

	DECLARE_LIBID(LIBID_Tangram)

private:
	map<void*, IUnknown*>					m_mapObjects;
	map<CString, ICreator*>					m_mViewFactory;
	map<DWORD,TangramThreadInfo*>			m_mapThreadInfo;	
	CMapStringToPtr							m_TabWndClassInfoDictionary;

	//.NET Version 4: 
	ICLRRuntimeHost*						m_pClrHost;

	BOOL IsWow64bit();
	BOOL Is64bitSystem();
	//BOOL _CheckNode(CTangramXmlParse* pXmlNode);
	CString _GetNames(CWndNode* pNode);
	void _clearObjects()
	{
		map<void*,IUnknown*>::iterator it;
		while((it = m_mapObjects.begin()) != m_mapObjects.end())
		{
			IUnknown* pUnknown = it->second;
			CCommonFunction::ClearObject<IUnknown>(pUnknown);
			m_mapObjects.erase(it);
		}
	}

	void GetJVMInfo();
    int ExitInstance();
	BOOL InitInstance();
	void AttachNode(void* pNodeEvents);
	void AddClsInfo(CString m_strObjID, int nType, CRuntimeClass* pClsInfo);
	void OnEvent(IEventProxy* pEvent, IDispatch* pCtrlDisp, IDispatch* pArgDisp);
	static LRESULT CALLBACK GetMessageProc(int nCode, WPARAM wParam, LPARAM lParam);
};

extern CTangramApp theApp;
