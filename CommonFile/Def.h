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
*
********************************************************************************/

#pragma once

#include <map>
#include "mso.h"
#include "Tangram.h"
#include "TangramXmlParse.h"
using namespace std;
using namespace ATL;

#pragma comment(lib, "imagehlp.lib")

typedef map<HWND, CString> CTaskPaneXmlMap;

#define TANGRAM_OBJECT_ENTRY_AUTO(clsid, class) \
    __declspec(selectany) ATL::_ATL_OBJMAP_CACHE __objCache__##class = { NULL, 0 }; \
	const ATL::_ATL_OBJMAP_ENTRY_EX __objMap_##class = {&clsid, class::UpdateRegistry, class::_ClassFactoryCreatorClass::CreateInstance, class::CreateInstance, &__objCache__##class, class::GetObjectDescription, class::GetCategoryMap, class::ObjectMain }; \
	extern "C" __declspec(allocate("ATL$__m")) __declspec(selectany) const ATL::_ATL_OBJMAP_ENTRY_EX* const __pobjMap_##class = &__objMap_##class; \
	OBJECT_ENTRY_PRAGMA(class)


namespace TangramCommon
{
	typedef struct tagCtrlInfo
	{
		HWND			m_hWnd;
		CString			m_strName;
		IWndPage*		m_pPage;
		IWndNode*		m_pNode;
		IDispatch*		m_pCtrlDisp;
	}CtrlInfo;

	typedef struct tagUCMAMSGInfo
	{
		CString			m_strData;
		CString			m_strMsg;
	}UCMAMSGInfo;

	class CTangramProxy
	{
	public:
		virtual void WindowCreated(LPCTSTR strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd) {};
		virtual void WindowDestroy(HWND hWnd) {};
	};

	class CTangramProxyBase
	{
	public:
		virtual void AttachNode(void* pNodeEvents){};
		virtual void OnEvent(IEventProxy* pEvent, IDispatch* pCtrlDisp, IDispatch* pArgDisp){};
	};

	class CTangramPackageProxy
	{
	public:
		virtual void CreateTangramToolWnd(){};
	};

	class CApplicationCLRProxyImpl
	{
	public:
		CApplicationCLRProxyImpl()
		{
			m_pProxy = NULL;
			m_strObjTypeName = _T("");
			m_strCollaborationScript = _T("");
		};

		CString			   m_strObjTypeName;
		CString			   m_strCollaborationScript;
		CTangramProxyBase* m_pProxy;
		virtual BSTR AttachObjEvent(IDispatch* EventObj, IDispatch* SourceObj, WindowEventType EventType, IDispatch* HTMLWindow) = 0;
		virtual HRESULT ActiveCLRMethod(BSTR bstrObjID, BSTR bstrMethod, BSTR bstrParam, BSTR bstrData) = 0;
		virtual IDispatch* CreateCLRObj(BSTR bstrObjID) = 0;
		virtual HRESULT ProcessCtrlMsg(HWND hCtrl, bool bShiftKey) = 0;
		virtual BOOL ProcessFormMsg(HWND hFormWnd, LONG lpMsg, int nMouseButton) = 0;
		virtual HRESULT CloseForms()=0;
		virtual IDispatch* TangramCreateObject(BSTR bstrObjID, long hParent, IWndNode* pHostNode)=0;
		virtual BOOL IsWinForm(HWND hWnd) = 0;
		virtual IDispatch* GetCLRControl(IDispatch* CtrlDisp, BSTR bstrNames)=0;
		virtual BSTR GetCtrlName(IDispatch* pCtrl)=0;
		virtual HWND GetMDIClientHandle(IDispatch* pMDICtrl)=0;
		virtual IDispatch* GetCtrlByName(IDispatch* CtrlDisp, BSTR bstrName, bool bFindInChild)=0;
		virtual HWND GetCtrlHandle(IDispatch* pCtrl)=0;
		virtual HWND IsCtrlCanNavigate(IDispatch* ctrl)=0;
		virtual void TangramAction(BSTR bstrXml, IWndNode* pNode)=0;
		virtual BSTR GetCtrlValueByName(IDispatch* CtrlDisp, BSTR bstrName, bool bFindInChild)=0;
		virtual CString GetCtrlInfo(long hWnd)=0;
		virtual HWND OnInitialDesigner(long hWnd, CString& strName)=0;
		virtual CString GetControlStrs(CString strLib) { return _T(""); };
		virtual void SelectNode(IWndNode* ) { };
		virtual void AttachPropertyGridCtrl(HWND hWnd ) { };
	};

	#define WM_TANGRAM_WEBNODEDOCCOMPLETE	(WM_USER + 0x00004001)
	#define WM_DESTROYTANGRAM				(WM_USER + 0x00004002)
	#define WM_SETHELPWND					(WM_USER + 0x00004003)
	#define WM_SETHELPWNDEX					(WM_USER + 0x00004004)
	#define WM_TABCHANGE					(WM_USER + 0x00004005)
	#define WM_TANGRAMMSG					(WM_USER + 0x00004006)
	#define WM_NAVIXTML						(WM_USER + 0x00004007)
	#define WM_GETTANGRAM					(WM_USER + 0x00004008)
	#define WM_MDICHILDMIN					(WM_USER + 0x00004009)
	#define WM_ACTIVEAPPXMLDOC				(WM_USER + 0x0000400a)
	#define WM_GETTANGRAMWINDOW				(WM_USER + 0x00006563)
	#define WM_OPENDOCUMENT					(WM_USER + 0x0000400b)
	#define WM_TANGRAMDATA					(WM_USER + 0x0000400c)
	#define WM_DOWNLOAD_MSG					(WM_USER + 0x0000400d)
	#define WM_VBAFORMCREATED				(WM_USER + 0x0000400e)
	#define WM_TANGRAMUCMAMSG				(WM_USER + 0x0000400f)
	#define WM_SWTCOMPONENTNOTIFY			(WM_USER + 0x00004010)
	#define WM_TANGRAMRESIZETNOTIFY			(WM_USER + 0x00004011)
	#define WM_RUNTIMEVBAFORMCREATED		(WM_USER + 0x00004012)
	#define WM_TANGRAMDESIGNMSG				(WM_USER + 0x00004013)
	#define WM_INSERTTREENODE				(WM_USER + 0x00004014)
	#define WM_REFRESHDATA					(WM_USER + 0x00004015)
	#define WM_GETSELECTEDNODEINFO			(WM_USER + 0x00004016)
	#define WM_TANGRAMDESIGNERCMD			(WM_USER + 0x00004017)
	#define WM_TANGRAMGETTREEINFO			(WM_USER + 0x00004018)
	#define WM_TANGRAMGETNODE				(WM_USER + 0x00004019)
	#define WM_TANGRAMUPDATENODE			(WM_USER + 0x0000401a)
	#define WM_TANGRAMGETDESIGNWND			(WM_USER + 0x0000401b)
	#define WM_TANGRAMISDOCUMENT			(WM_USER + 0x0000401c)
	#define WM_TGM_SETACTIVEPAGE			(WM_USER + 0x0000401d)
	#define WM_TGM_SET_CAPTION				(WM_USER + 0x0000401e)	
	#define WM_GETNODEINFO					(WM_USER + 0x0000401f)
	#define WM_CREATETABPAGE				(WM_USER + 0x00004020)
	#define WM_ACTIVETABPAGE				(WM_USER + 0x00004021)
	#define WM_MODIFYTABPAGE				(WM_USER + 0x00004022)
	#define WM_ADDTABPAGE					(WM_USER + 0x00004023)
	#define WM_UPDATETANGRAMOBJ				(WM_USER + 0x00004024)
	#define WM_TANGRAMNEWOUTLOOKOBJ			(WM_USER + 0x00004025)
	#define WM_TANGRAMACTIVEINSPECTORPAGE	(WM_USER + 0x00004026)
	#define WM_TANGRAMITEMLOAD				(WM_USER + 0x00003027)
	#define WM_TANGRAMSAVE					(WM_USER + 0x00004028)
	#define WM_USER_TANGRAMTASK				(WM_USER + 0x00004029)
	#define WM_UPLOADFILE					(WM_USER + 0x00004030)
	#define WM_REMOVERESTKEY				(WM_USER + 0x00004031)
	#define WM_INITPROPERTYGRID				(WM_USER + 0x00004032)
	#define WM_NAMECHANGED					(WM_USER + 0x00004033)
	#define WM_SETWNDFOCUSE					(WM_USER + 0x00004034)
	#define WM_OFFICEOBJECTCREATED			(WM_USER + 0x00004035)
	#define WM_TANGRAMECLIPSEINFO			(WM_USER + 0x00004036)

	typedef enum ViewType
	{
		BlankView = 0x00000001,
		ActiveX = 0x00000002,
		Splitter = 0x00000004,
		TabbedWnd = 0x00000008,

		CLRCtrl = 0x00000010,
		CLRForm = 0x00000020,
		CLRWnd = 0x00000040,
		TangramView = 0x00000080,
		TreeView = 0x000000b0
	}ViewType;

	#define TGM_NAME				_T("name")
	#define TGM_CAPTION				_T("caption")
	#define TGM_NODE_TYPE			_T("id")
	#define TGM_CNN_ID				_T("cnnid")
	#define TGM_HEIGHT				_T("height")
	#define TGM_WIDTH				_T("width")
	#define TGM_STYLE				_T("style")
	#define TGM_ACTIVE_PAGE			_T("activepage")
	#define TGM_TAG					_T("tag")
	#define TGM_NODE				_T("node")

	#define TGM_ROWS				_T("rows")
	#define TGM_COLS				_T("cols")


	#define TGM_SPLITTER			_T("splitter")
	#define TGM_TABBED				_T("tab")

	#define TGM_SETTING_HEAD		_T("#$^&TANGRAM")
	#define TGM_SETTING_FOMRAT		_T("#$^&TANGRAM[%ld][%ld]")

	#define TGM_S_EXCEL_INPUT		1
};

using namespace TangramCommon;