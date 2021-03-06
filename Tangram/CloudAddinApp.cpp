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

// TangramApp.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "CloudAddinApp.h" 
#include "WndPage.h"
#include "WebTaskObj.h"
#include "TangramTreeView.h"
#include "TangramHtmlTreeWnd.h"
#include "TangramHtmlTreeExWnd.h"
#include "ProgressFX.h"
#include "HourglassFX.h"
#include "atlenc.h"
#include <string>
#include <jni.h>
#include <VersionHelpers.h> 

#include "fm20.h"
#include "CloudUtilities\CloudAddinDownLoad.h"
#include "VisualStudioPlus\CloudAddinVSAddin.h"
#include "EclipsePlus\EclipseAddin.h"

#include "OfficePlus\WordPlus\WordAddin.h"
#include "OfficePlus\WordPlus\WordPlusDoc.h"
#include "OfficePlus\LyncPlus\LyncAddin.h"
#include "OfficePlus\ExcelPlus\ExcelAddin.h"
#include "OfficePlus\ExcelPlus\ExcelPlusWnd.h"
#include "OfficePlus\VisioPlus\VisioAddin.h"
#include "OfficePlus\AccessPlus\AccessAddin.h"
#include "OfficePlus\OutLookPlus\OutLookAddin.h"
#include "OfficePlus\ProjectPlus\ProjectAddin.h"
#include "OfficePlus\PowerPointPlus\PowerPointAddin.h"

#include "NodeWnd.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "TabbedView.h"
#include "SplitterWnd.h"
#include "TangramCoreEvents.h"
#include "TangramWinRTProxy.h"
#include "TangramWinRTProxy.c"
//#include "TangramEclipse_i.h"

using namespace OfficeCloudPlus;
using namespace OfficeCloudPlus::WordPlus;
using namespace OfficeCloudPlus::ExcelPlus;

void CHelperWnd::OnFinalMessage(HWND hWnd)
{
	CWindowImpl::OnFinalMessage(hWnd);
	if (m_strID != _T(""))
	{
		auto it = theApp.m_mapRemoteTangramCore.find(m_strID);
		if (it != theApp.m_mapRemoteTangramCore.end())
		{
			theApp.m_mapRemoteTangramCore.erase(m_strID);
		}
		auto it2 = theApp.m_pHostCore->m_mapAppDispDic.find(m_strID);
		if (it2 != theApp.m_pHostCore->m_mapAppDispDic.end())
		{
			theApp.m_pHostCore->m_mapAppDispDic.erase(it2);
		}
	}
	delete this;
}

// Description  : the unique App object
CTangramApp theApp;

CTangramApp::CTangramApp()
{
	m_dwThreadID				= 0;
	m_nAppID					= -1;
	m_bOfficeApp				= false;
	m_bMSAddin					= false;
	m_bInitTangramJava			= false;
	m_bClose					= false;
	m_bCLRStart					= false;
	m_b32Process				= false;
	m_bWinFormActived			= false;
	m_bEnableProcessFormTabKey	= false;
	m_hCBTHook					= NULL;
	m_hActiveWnd				= NULL;
	m_hModule					= nullptr;
	m_pTangramProxy				= nullptr;
	m_pAppDisp					= nullptr;
	m_pPage						= nullptr;
	m_pClrHost					= nullptr;
	m_pHostCore					= nullptr;
	m_pEventProxy				= nullptr;
	m_pDocDOMTree				= nullptr;
	m_pWndNode					= nullptr;
	m_pTangram					= nullptr;
	m_pTangramAddin				= nullptr;
	m_pJVMenv					= nullptr;
	m_pICollaborationProxy		= nullptr;
	m_pJavaProxy				= nullptr;
	m_pWndFrame					= nullptr;
	m_pDesignerFrame			= nullptr;
	m_pDesignerWndPage			= nullptr;
	m_pCloudAddinCLRProxy		= nullptr;
	m_pCurWorkNodeParse			= nullptr;
	m_pDesignWindowNode			= nullptr;
	m_pMDIClientBKWnd			= nullptr;
	m_pDesigningFrame			= nullptr;
	m_pHostDesignUINode			= nullptr;
	m_pTangramPackageProxy		= nullptr;
	m_varApplication.vt			= VT_EMPTY;
	m_strExeName				= _T("");
	m_strConfigFile				= _T("");
	m_strCurrentKey				= _T("");
	m_strTangramCLRPath			= _T("");
#ifdef _DEBUG
	m_nTangram					= 0;
	m_nTangramObj				= 0;
	m_nTangramFrame				= 0;
	m_nOfficeDocs			= 0;
	m_nOfficeDocsSheet		= 0;
#endif
	m_pJVM = nullptr;
}

CTangramApp::~CTangramApp()
{
	DWORD dw = 0;
	if (m_pHostCore)
	{
		if (m_pEventTypeInfo)
		{
			ITypeInfo* pDisp = m_pEventTypeInfo.Detach();
			dw = pDisp->Release();
		}

	#ifdef _DEBUG
		if(m_nTangram)
			TRACE(_T("Tangram Count: %d\n"), m_nTangram);
		if(m_nTangramObj)
			TRACE(_T("TangramObj Count: %d\n"), m_nTangramObj);
		if(m_nTangramFrame)
			TRACE(_T("TangramFrame Count: %d\n"), m_nTangramFrame);
		if(m_nOfficeDocs)
			TRACE(_T("TangramOfficeDoc Count: %d\n"), m_nOfficeDocs);
		if(m_nOfficeDocsSheet)
			TRACE(_T("TangramExcelWorkBookSheet Count: %d\n"), m_nOfficeDocsSheet);
	#endif
		delete m_pHostCore;
	}

	TRACE(_T("TangramApp Terminate: %x\n"), this);

	m_pTangram = nullptr;

	if (m_pClrHost)
	{
		if (m_nAppID == -1)
		{
			if (m_bCLRStart == false)
			{
				TRACE(_T("------------------Begin Stop CLR------------------------\n"));
				TRACE(_T("------------------Stop CLR Successed!------------------------\n"));
				HRESULT hr;
				hr = m_pClrHost->Stop();
				ASSERT(hr == S_OK);
				dw = m_pClrHost->Release();
				ASSERT(dw == 0);
				TRACE(_T("------------------ClrHost Release Successed!------------------------\n"));
				TRACE(_T("------------------End Stop CLR------------------------\n"));
			}
		}
	}
}

BOOL CTangramApp::IsUserAdministrator()
{
	BOOL bRet = false;
	PSID psidRidGroup;
	SID_IDENTIFIER_AUTHORITY siaNtAuthority = SECURITY_NT_AUTHORITY;

	bRet = AllocateAndInitializeSid(&siaNtAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, &psidRidGroup);
	if (bRet)
	{
		if (!CheckTokenMembership(NULL, psidRidGroup, &bRet))
			bRet = false;
		FreeSid(psidRidGroup);
	}

	return (BOOL)bRet;
}

bool CTangramApp::CheckUrl(CString&   url)
{
	char*		res = nullptr;
	char		dwCode[20];
	DWORD		dwIndex, dwCodeLen;
	HINTERNET   hSession, hFile;

	url.MakeLower();

	hSession = InternetOpen(_T("Tangram"), INTERNET_OPEN_TYPE_PRECONFIG, NULL, NULL, 0);
	if (hSession)
	{
		hFile = InternetOpenUrl(hSession, url, NULL, 0, INTERNET_FLAG_RELOAD, 0);
		if (hFile == NULL)
		{
			InternetCloseHandle(hSession);
			return false;
		}
		dwIndex = 0;
		dwCodeLen = 10;
		HttpQueryInfo(hFile, HTTP_QUERY_STATUS_CODE, dwCode, &dwCodeLen, &dwIndex);
		res = dwCode;
		if (strcmp(res, "200 ") || strcmp(res, "302 "))
		{
			//200,302未重定位标志    
			if (hFile)
				InternetCloseHandle(hFile);
			InternetCloseHandle(hSession);
			return   true;
		}
	}
	return   false;
}

DWORD CTangramApp::ExecCmd(const CString cmd, const BOOL setCurrentDirectory)
{
	BOOL  bReturnVal = false;
	STARTUPINFO  si;
	DWORD  dwExitCode = ERROR_NOT_SUPPORTED;
	SECURITY_ATTRIBUTES saProcess, saThread;
	PROCESS_INFORMATION process_info;

	ZeroMemory(&si, sizeof(si));
	si.cb = sizeof(si);

	saProcess.nLength = sizeof(saProcess);
	saProcess.lpSecurityDescriptor = NULL;
	saProcess.bInheritHandle = true;

	saThread.nLength = sizeof(saThread);
	saThread.lpSecurityDescriptor = NULL;
	saThread.bInheritHandle = false;

	CString currentDirectory = _T("");

	bReturnVal = CreateProcess(NULL,
		(LPTSTR)(LPCTSTR)cmd,
		&saProcess,
		&saThread,
		false,
		DETACHED_PROCESS,
		NULL,
		currentDirectory,
		&si,
		&process_info);

	if (bReturnVal)
	{
		CloseHandle(process_info.hThread);
		WaitForSingleObject(process_info.hProcess, INFINITE);
		GetExitCodeProcess(process_info.hProcess, &dwExitCode);
		CloseHandle(process_info.hProcess);
	}
	//else
	//{
	//	DWORD dw =  GetLastError();
	//	dwExitCode = dw;
	//}

	return dwExitCode;
}

void CTangramApp::AttachNode(void* pNodeEvents)
{
	CWndNodeEvents*	m_pCLREventConnector = (CWndNodeEvents*)pNodeEvents;
	CWndNode* pNode = (CWndNode*)m_pCLREventConnector->m_pWndNode;
	pNode->m_pCLREventConnector = m_pCLREventConnector;
};

void CTangramApp::OnEvent(IEventProxy* pEvent, IDispatch* pCtrlDisp, IDispatch* pArgDisp)
{
	CEventProxy* pTangramEvent = (CEventProxy*)pEvent;
	if (pTangramEvent)
	{
		IDispatch * pConnection = static_cast<IDispatch *>(pTangramEvent->m_vec.GetAt(0));

		if (pConnection)
		{
			CComVariant avarParams[2];
			avarParams[1] = pCtrlDisp;
			avarParams[1].vt = VT_DISPATCH;
			avarParams[0] = pArgDisp;
			avarParams[0].vt = VT_DISPATCH;
			CComVariant varResult;
			DISPPARAMS params = { avarParams, NULL, 2, 0 };
			pConnection->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, nullptr, nullptr);
		}
	}
};

BOOL CTangramApp::LoadImageFromResource(ATL::CImage *pImage, HMODULE hMod, UINT nResID, LPCTSTR lpTyp)
{
	if (pImage == nullptr)
		return false;

	pImage->Destroy();

	// 查找资源
	HRSRC hRsrc = ::FindResource(hMod, MAKEINTRESOURCE(nResID), lpTyp);
	if (hRsrc == NULL)
		return false;
	HGLOBAL hImgData = ::LoadResource(hMod, hRsrc);
	if (hImgData == NULL)
	{
		::FreeResource(hImgData);
		return false;
	}

	// 锁定内存中的指定资源
	LPVOID lpVoid = ::LockResource(hImgData);

	LPSTREAM pStream = nullptr;
	DWORD dwSize = ::SizeofResource(hMod, hRsrc);
	HGLOBAL hNew = ::GlobalAlloc(GHND, dwSize);
	LPBYTE lpByte = (LPBYTE)::GlobalLock(hNew);
	::memcpy(lpByte, lpVoid, dwSize);

	// 解除内存中的指定资源
	::GlobalUnlock(hNew);
	// 从指定内存创建流对象
	HRESULT ht = ::CreateStreamOnHGlobal(hNew, true, &pStream);
	if (ht == S_OK)
	{
		// 加载图片
		pImage->Load(pStream);

	}
	GlobalFree(hNew);
	// 释放资源
	::FreeResource(hImgData);
	return true;
}

TangramThreadInfo* CTangramApp::GetThreadInfo(DWORD ThreadID)
{
	TangramThreadInfo* pInfo = nullptr;

	DWORD nThreadID = ThreadID;
	if (nThreadID == 0)
		nThreadID = GetCurrentThreadId();
	map<DWORD, TangramThreadInfo*>::iterator it = m_mapThreadInfo.find(nThreadID);
	if (it != m_mapThreadInfo.end())
	{
		pInfo = it->second;
	}
	else
	{
		pInfo = new TangramThreadInfo();
		pInfo->m_hGetMessageHook = NULL;
		m_mapThreadInfo[nThreadID] = pInfo;
	}
	return pInfo;
}

int Base64Encode(PBYTE pSrc, LPSTR pDst, int nSrcLen, int nMaxLineLen)
{
	const char EnBase64Tab[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
	unsigned char c1, c2, c3;
	int nDstLen = 0;

	int nLineLen = 0;
	int nDiv = nSrcLen / 3;
	int nMod = nSrcLen % 3;

	int i;

	for (i = 0; i < nDiv; i++)
	{
		c1 = *pSrc++;
		c2 = *pSrc++;
		c3 = *pSrc++;

		*pDst++ = EnBase64Tab[c1 >> 2];
		*pDst++ = EnBase64Tab[((c1 << 4) | (c2 >> 4)) & 0x3f];
		*pDst++ = EnBase64Tab[((c2 << 2) | (c3 >> 6)) & 0x3f];
		*pDst++ = EnBase64Tab[c3 & 0x3f];
		nLineLen += 4;
		nDstLen += 4;

		if (nLineLen >= nMaxLineLen - 4)
		{
			*pDst++ = '\r';
			*pDst++ = '\n';
			nLineLen = 0;
			nDstLen += 2;
		}
	}

	if (nMod == 1)
	{
		c1 = *pSrc++;
		*pDst++ = EnBase64Tab[(c1 & 0xfc) >> 2];
		*pDst++ = EnBase64Tab[((c1 & 0x03) << 4)];
		*pDst++ = '=';
		*pDst++ = '=';
		nLineLen += 4;
		nDstLen += 4;
	}
	else if (nMod == 2)
	{
		c1 = *pSrc++;
		c2 = *pSrc++;
		*pDst++ = EnBase64Tab[(c1 & 0xfc) >> 2];
		*pDst++ = EnBase64Tab[((c1 & 0x03) << 4) | ((c2 & 0xf0) >> 4)];
		*pDst++ = EnBase64Tab[((c2 & 0x0f) << 2)];
		*pDst++ = '=';
		nDstLen += 4;
	}

	*pDst = '\0';
	return nDstLen;
}

BOOL CTangramApp::InitInstance()
{
	::GetTempPath(MAX_PATH, m_szBuffer);
	theApp.m_strTempPath = CString(m_szBuffer);

	TCHAR szDriver[MAX_PATH] = { 0 };
	TCHAR szDir[MAX_PATH] = { 0 };
	TCHAR szExt[MAX_PATH] = { 0 };
	TCHAR szFile2[MAX_PATH] = { 0 };
	::GetModuleFileName(NULL, m_szBuffer, MAX_PATH);
	CString strFile = CString(m_szBuffer);
	strFile += _T(".tangram");
	_tsplitpath_s(m_szBuffer, szDriver, szDir, szFile2, szExt);
	m_strAppPath = szDriver;
	m_strAppPath += szDir;
	CString path(m_szBuffer);
	int nPos = path.ReverseFind('\\');
	CString strName = path.Mid(nPos + 1);
	nPos = strName.Find(_T("."));
	m_strExeName = strName.Left(nPos);
	m_strExeName.MakeLower();
	m_strStartJar = _T("");
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);
	memset(m_szBuffer, 0, sizeof(m_szBuffer));
	m_dwThreadID = ::GetCurrentThreadId();
	theApp.m_mapValInfo[_T("AppPath")] = CComVariant(m_strAppPath);

	SHGetFolderPath(NULL, CSIDL_WINDOWS, NULL, 0, m_szBuffer);
	m_strTangramCLRPath = CString(m_szBuffer) + _T("\\Microsoft.NET\\assembly\\");
	m_hModule = ::GetModuleHandle(_T("tangram.dll"));
	//regsvr32
	// get arguments
//	if (m_strExeName.CompareNoCase(_T("regsvr32")))
//	{
//		CefMainArgs main_args(GetModuleHandle(NULL));
//
//		// Execute the secondary process, if any.
//		int exit_code = CefExecuteProcess(main_args, m_cefApp.get(), NULL);
//		if (exit_code >= 0)
//			return exit_code;
//
//		// setup settings
//		CString szCEFCache;
//		CString szPath;
//		INT nLen = GetTempPath(0, NULL) + 1;
//		GetTempPath(nLen, szPath.GetBuffer(nLen));
//
//		// save path
//		szCEFCache.Format(_T("%scache\0\0"), szPath);
//
//		// set settings
//		CefSettings settings;
//
//		settings.multi_threaded_message_loop = FALSE;
//		CefString(&settings.cache_path) = szCEFCache;
//
//		void* sandbox_info = NULL;
//#if CEF_ENABLE_SANDBOX
//		// Manage the life span of the sandbox information object. This is necessary
//		// for sandbox support on Windows. See cef_sandbox_win.h for complete details.
//		CefScopedSandboxInfo scoped_sandbox;
//		sandbox_info = scoped_sandbox.sandbox_info();
//#else
//		settings.no_sandbox = TRUE;
//#endif
//
//		//CEF Initiaized
//		m_bCEFInitialized = CefInitialize(main_args, settings, m_cefApp.get(), sandbox_info);
//
//
//	}

	m_b32Process = !(Is64bitSystem() && !IsWow64bit());
	if (m_b32Process)
		m_strTangramCLRPath += _T("GAC_32\\TangramCLR\\v4.0_1.0.1992.1963__1bcc94f26a4807a7\\TangramCLR.Dll");
	else
		m_strTangramCLRPath += _T("GAC_64\\TangramCLR\\v4.0_1.0.1992.1963__1bcc94f26a4807a7\\TangramCLR.Dll");

	if (!::PathFileExists(m_strTangramCLRPath))
	{
		m_strTangramCLRPath = _T("TangramCLR.dll");
		if (::PathFileExists(m_strAppPath + m_strTangramCLRPath) == false)
		{
			CString strPath = m_strProgramFilePath;
			strPath += _T("\\tangram\\tangramclr.dll");
			if (::PathFileExists(strPath))
				m_strTangramCLRPath = strPath;
			else
				m_strTangramCLRPath = _T("");
		}
		else
			m_strTangramCLRPath = m_strAppPath + m_strTangramCLRPath;
	}

	HRESULT hr = SHGetFolderPath(NULL, CSIDL_APPDATA, NULL, 0, m_szBuffer);
	m_strAppDataPath = CString(m_szBuffer);
	m_strAppDataPath += _T("\\TangramData\\");

	hr = SHGetFolderPath(NULL, CSIDL_PROGRAM_FILES, NULL, 0, m_szBuffer);
	m_strProgramFilePath = CString(m_szBuffer);

	m_strAppDataPath += m_strExeName;
	m_strAppDataPath += _T("\\"); 

	WNDCLASS wndClass;
	wndClass.style = CS_DBLCLKS;
	wndClass.lpfnWndProc = ::DefWindowProc;
	wndClass.cbClsExtra = 0;
	wndClass.cbWndExtra = 0;
	wndClass.hInstance = AfxGetInstanceHandle();
	wndClass.hIcon = 0;
	wndClass.hCursor = ::LoadCursor(NULL, IDC_ARROW);
	wndClass.hbrBackground = 0;
	wndClass.lpszMenuName = NULL;
	wndClass.lpszClassName = _T("Tangram Splitter Class");

	RegisterClass(&wndClass);

	m_lpszSplitterClass = wndClass.lpszClassName;

	wndClass.lpfnWndProc = TangramWndProc;
	wndClass.style = CS_HREDRAW | CS_VREDRAW;
	wndClass.lpszClassName = L"Tangram Window Class";

	RegisterClass(&wndClass);

	AddClsInfo(_T("HostView"), TangramView, RUNTIME_CLASS(CNodeWnd));
	AddClsInfo(TGM_SPLITTER, Splitter, RUNTIME_CLASS(CSplitterNodeWnd));
	AddClsInfo(TGM_TABBED, TabbedWnd, RUNTIME_CLASS(CTabbedNodeView));

	CTangramXmlParse m_Parse;
	if (m_Parse.LoadFile(strFile))
	{
		m_strConfigFile = m_Parse.xml();
		m_strStartJar = m_Parse.attr(_T("startjar"), _T(""));
		CTangramXmlParse* pChild = m_Parse.GetChild(m_strExeName);
		if (pChild)
		{
			int nCount = pChild->GetCount();
			if (nCount)
			{
				CTangramXmlParse* pParse = nullptr;
				for (int i = 0; i < nCount; i++)
				{
					pParse = pChild->GetChild(i);
					m_mapValInfo[pParse->name().MakeLower()] = CComVariant(pParse->xml());
				}
			}
		}

		pChild = m_Parse.GetChild(_T("Collaboration"));
		if (pChild)
		{
			theApp.m_mapValInfo[_T("collaborationscript")] = CComVariant(pChild->xml());
		}
	}

	BOOL bOfficeApp = false;
	HMODULE hModule = ::GetModuleHandle(_T("mso.dll"));
	if (hModule)
	{
		bOfficeApp = true;
		m_mapOfficeExeID[_T("winword")] = 0;
		m_mapOfficeExeID[_T("excel")] = 1;
		m_mapOfficeExeID[_T("powerpnt")] = 2;
		m_mapOfficeExeID[_T("outlook")] = 3;
		m_mapOfficeExeID[_T("msaccess")] = 4;
		//m_mapOfficeExeID[_T("infopath")] = 5;
		m_mapOfficeExeID[_T("project")] = 6;
		//m_mapOfficeExeID[_T("onenote")] = 7;
		m_mapOfficeExeID[_T("visio")] = 8;
		m_mapOfficeExeID[_T("lync")] = 9;
	}
	else if(::GetModuleHandle(_T("kso.dll")))
	{
		bOfficeApp = true;
		m_mapOfficeExeID[_T("wps")] = 0;
		m_mapOfficeExeID[_T("et")] = 1;
		m_mapOfficeExeID[_T("wpp")] = 2;
	}
	if (bOfficeApp)
	{
		//::FreeLibrary(hModule);
	}
	else
	{
		hModule = ::GetModuleHandle(_T("tangrameclipse.dll"));
		if (hModule)
		{
			m_pTangram = new CComObject < EclipseCloudPlus::EclipsePlus::CEclipseCloudAddin >;
		}
		else
		{
#ifndef _WIN64
			if (m_strExeName == _T("devenv"))
			{
				CString strVer = theApp.GetFileVer();
				int nPos = strVer.Find(_T("."));
				int nVer = _wtoi(strVer.Left(nPos));
				if (nVer > 9)
				{
					m_pTangram = new CComObject < VSCloudPlus::VisualStudioPlus::CVSCloudAddin >;
				}
				else
				{
					m_pTangram = new CComObject < CTangram >;
				}
			}
			else
			{
				m_pTangram = new CComObject < CTangram >;
			}
#else
			m_pTangram = new CComObject < CTangram >;
#endif	
			//InitEclipse();
		}
	}
	//CString strAddinID = m_strExeName;
	//strAddinID += _T(".TangramAddin");

	//CComPtr<ITangramAddin> pTangramAddin;
	//pTangramAddin.CoCreateInstance(strAddinID.AllocSysString());
	//if (pTangramAddin)
	//{
	//	m_pTangramAddin = pTangramAddin.Detach();
	//	m_pTangramAddin->AddRef();
	//}
	if(m_pHostCore)
		m_pHostCore->Init();
	//return true;

	return true;
}

void CTangramApp::GetJVMInfo()
{

}

int CTangramApp::ExitInstance()
{
	if (m_pHostCore)
	{
		auto it = m_pHostCore->m_mapWebTask.begin();
		while (it != m_pHostCore->m_mapWebTask.end())
		{
			if (it->second->m_hHandle == 0)
				delete it->second;
			else
			{
				::WaitForSingleObject(it->second->m_hHandle, INFINITE);
				::PostThreadMessage(it->second->m_dwThreadID, WM_QUIT, 0, 0);
			}
			it = m_pHostCore->m_mapWebTask.begin();
		}

		if (::IsWindow(m_pHostCore->m_hHostWnd))
		{
			::DestroyWindow(m_pHostCore->m_hHostWnd);
		}

		if (m_pTangramAddin != nullptr)
		{
			DWORD dw = m_pTangramAddin->Release();
			while (dw > 0)
				dw = m_pTangramAddin->Release();
		}
		//if (m_pJavaProxy)
		//{
		//	DWORD dw = m_pJavaProxy->Release();
		//	while (dw)
		//	{
		//		dw = m_pJavaProxy->Release();
		//	}
		//	m_pJavaProxy = nullptr;
		//}
	}

	m_bClose = true;
	TRACE(_T("Begin Tangram ExitInstance :%p\n"), this);

	if (m_hCBTHook)
		UnhookWindowsHookEx(m_hCBTHook);

	for (auto it : m_mapThreadInfo)
	{
		if (it.second->m_hGetMessageHook)
		{
			UnhookWindowsHookEx(it.second->m_hGetMessageHook);
			it.second->m_hGetMessageHook = NULL;
		}
		delete it.second;
	}
	m_mapThreadInfo.erase(m_mapThreadInfo.begin(), m_mapThreadInfo.end());

	_clearObjects();
	auto it2 = theApp.m_mapWindowPage.begin();
	while (it2 != theApp.m_mapWindowPage.end())
	{
		delete it2->second;
		//theApp.m_mapWindowPage.erase(it2);
		it2 = theApp.m_mapWindowPage.begin();
	}

	CString strIndex = _T("");
	void* pObj;
	for (POSITION pos = m_TabWndClassInfoDictionary.GetStartPosition(); pos != NULL; )
	{
		m_TabWndClassInfoDictionary.GetNextAssoc(pos, strIndex, (void*&)pObj);
		delete pObj;
	}
	m_TabWndClassInfoDictionary.RemoveAll();
	TRACE(_T("End Tangram ExitInstance :%p\n"), this);

	//if (m_bCEFInitialized) {
	//	// closing stop work loop
	//	m_bCEFInitialized = FALSE;
	//	// release CEF app
	//	m_cefApp = NULL;
	//	// clean up CEF loop
	//	CefDoMessageLoopWork();
	//	// shutdown CEF
	//	CefShutdown();
	//}
		
	return CWinApp::ExitInstance();
}

BOOL CTangramApp::Create(CWndNode* pWindowNode, LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext)
{
	//CComPtr<ITangramWinRT> pRT;
	//pRT.CoCreateInstance(CLSID_TangramWinRT);
	//pRT.CoCreateInstance(L"TangramWinRT.TangramWinRT");
	//pRT->CreateJSRuntime();
	HWND hWnd = 0;
	BOOL bRet = false;
	CString strKey = pWindowNode->m_strID;
	strKey.MakeLower().Trim();
	CNodeWnd* pTangramDesignView = (CNodeWnd*)pWindowNode->m_pHostWnd;

	ICreator* pViewFactoryDisp = nullptr;
	if (strKey.CompareNoCase(_T("ActiveX")) == 0 || strKey.CompareNoCase(_T("CLRCtrl")) == 0)
	{
		if (strKey.CompareNoCase(_T("CLRCtrl")) == 0 || pWindowNode->m_strCnnID.Find(_T(",")) != -1)
			LoadCLR();

		if (strKey.CompareNoCase(_T("CLRCtrl")) == 0)
			pWindowNode->m_nViewType = CLRCtrl;
		else if (strKey.CompareNoCase(_T("ActiveX")) == 0)
			pWindowNode->m_nViewType = ActiveX;
		if (pWindowNode->m_strCnnID.Find(_T("//")) == -1 && ::PathFileExists(pWindowNode->m_strCnnID) == false)
		{
			CString strPath = theApp.m_strAppPath + _T("TangramWebPage\\") + pWindowNode->m_strCnnID;
			if (::PathFileExists(strPath))
				pWindowNode->m_strCnnID = strPath;
		}

		hWnd = pWindowNode->CreateView(pParentWnd->m_hWnd, pWindowNode->m_strCnnID);

		CComBSTR bstrExtenderID(L"");
		pWindowNode->get_Attribute(_T("extender"), &bstrExtenderID);
		pWindowNode->m_strExtenderID = OLE2T(bstrExtenderID);
		pWindowNode->m_strExtenderID.Trim();
		if (pWindowNode->m_strExtenderID != _T(""))
		{
			CComPtr<IDispatch> pDisp;
			pDisp.CoCreateInstance(bstrExtenderID);
			if (pDisp)
			{
				pWindowNode->m_pExtender = pDisp.Detach();
				pWindowNode->m_pExtender->AddRef();
			}
		}

		pTangramDesignView->m_bCreateExternal = true;
		bRet = true;
	}
	else
	{
		strKey = pWindowNode->m_strCnnID;
		auto iter = m_mViewFactory.find(strKey);
		if (iter != m_mViewFactory.end())
			pViewFactoryDisp = iter->second;
		else
		{
			if (strKey.CompareNoCase(_T("hostapp")) == 0)
			{
				if (m_pAppDisp)
				{
					m_pAppDisp->QueryInterface(IID_ICreator, (void**)&pViewFactoryDisp);
					if (pViewFactoryDisp)
						m_mViewFactory[strKey] = pViewFactoryDisp;
				}
			}
			else
			{
				if (pWindowNode->m_strID.CompareNoCase(_T("TreeView")))
				{
					CComPtr<ICreator> pFactoryDisp;
					HRESULT hr = pFactoryDisp.CoCreateInstance(strKey.AllocSysString());
					if (hr == S_OK)
					{
						pViewFactoryDisp = pFactoryDisp.Detach();
						m_mViewFactory[strKey] = pViewFactoryDisp;
					}
				}
			}
		}

		if (pViewFactoryDisp)
		{
			long _nRet = 0;
			pViewFactoryDisp->Create(pParentWnd ? (long)pParentWnd->m_hWnd : 0, pWindowNode, &_nRet);
			hWnd = (HWND)_nRet;
			pWindowNode->m_nID = ::GetWindowLong(hWnd, GWL_ID);
		}
	}

	if (!::IsWindow(pTangramDesignView->m_hWnd) && hWnd&&pTangramDesignView->SubclassWindow(hWnd))
	{
		::SetWindowLong(hWnd, GWL_STYLE, dwStyle | WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS);
		::SetWindowLong(hWnd, GWL_ID, nID);

		pTangramDesignView->m_bCreateExternal = true;
		bRet = true;
	}

	if (hWnd == 0)
	{
		hWnd = CreateWindow(L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, pParentWnd->m_hWnd, (HMENU)nID, AfxGetInstanceHandle(), NULL);
	}
	if (::IsWindow(pTangramDesignView->m_hWnd) == false)
		bRet = pTangramDesignView->SubclassWindow(hWnd);

	if (pWindowNode->m_strID.CompareNoCase(_T("TreeView")) == 0)
	{
		CComBSTR bstrStyle(L"");
		pWindowNode->get_Attribute(CComBSTR(L"Style"), &bstrStyle);
		pWindowNode->m_nViewType = TreeView;
		CString _strStyle = OLE2T(bstrStyle);
		if (_strStyle != _T(""))
		{
			pTangramDesignView->m_pXHtmlTree = new CProgressFX< CHourglassFX< CTangramHtmlTreeEx2Wnd > >;
		}
		else
			pTangramDesignView->m_pXHtmlTree = new CTangramHtmlTreeWnd();
		pWindowNode->m_pDisp = pTangramDesignView->m_pXHtmlTree->m_pObj;
		pWindowNode->m_pDisp->AddRef();
		pTangramDesignView->m_pXHtmlTree->m_pHostWnd = pTangramDesignView;
		bRet = true;

		DWORD dwStyle = TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT |
			TVS_SHOWSELALWAYS | /*TVS_EDITLABELS |TVS_NOTOOLTIPS |*/
			WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP;

		CRect rect(0, 0, 100, 100);
		VERIFY(pTangramDesignView->m_pXHtmlTree->Create(dwStyle, rect, pTangramDesignView, 100));

		CComBSTR bstrCheckBox(L"");
		pWindowNode->get_Attribute(CComBSTR(L"CheckBoxes"), &bstrCheckBox);
		CString strCheckBox = OLE2T(bstrCheckBox);
		CComBSTR bstrSmartCheckBox(L"");
		pWindowNode->get_Attribute(CComBSTR(L"SmartCheckBox"), &bstrSmartCheckBox);
		CString strSmartCheckBox = OLE2T(bstrSmartCheckBox);
		CComBSTR bstrSetHtml(L"");
		pWindowNode->get_Attribute(CComBSTR(L"SetHtml"), &bstrSetHtml);
		CString strSetHtml = OLE2T(bstrSetHtml);
		CComBSTR bstrStripHtml(L"");
		pWindowNode->get_Attribute(CComBSTR(L"StripHtml"), &bstrStripHtml);
		CString strStripHtml = OLE2T(bstrStripHtml);
		CComBSTR bstrImages(L"");
		pWindowNode->get_Attribute(CComBSTR(L"Images"), &bstrImages);
		CString strImages = OLE2T(bstrImages);
		CComBSTR bstrReadOnly(L"");
		pWindowNode->get_Attribute(CComBSTR(L"ReadOnly"), &bstrReadOnly);
		CString strReadOnly = OLE2T(bstrReadOnly);

		int r = 0, g = 0, b = 0;
		CComBSTR bstrBKColor(L"");
		pWindowNode->get_Attribute(CComBSTR(L"BKColor"), &bstrBKColor);
		CString strBKColor = OLE2T(bstrBKColor);

		COLORREF colorBK = RGB(255, 255, 255);
		if (strBKColor != _T(""))
		{
			_stscanf_s(strBKColor, _T("RGB(%d,%d,%d)"), &r, &g, &b);
			colorBK = RGB(r, g, b);
		}

		CComBSTR bstrSeparatorColor(L"");
		pWindowNode->get_Attribute(CComBSTR(L"SeparatorColor"), &bstrSeparatorColor);
		CString strSeparatorColor = OLE2T(bstrSeparatorColor);

		COLORREF colorSeparator = RGB(0, 0, 255);
		if (bstrSeparatorColor != _T(""))
		{
			_stscanf_s(strSeparatorColor, _T("RGB(%d,%d,%d)"), &r, &g, &b);
			colorSeparator = RGB(r, g, b);
		}
		pTangramDesignView->m_pXHtmlTree->Initialize(strCheckBox != _T("") ? true : false, true)
			.SetSmartCheckBox(strSmartCheckBox != _T("") ? true : false)
			.SetHtml(true)
			//.SetHtml(strSetHtml != _T("") ? true : false)
			.SetStripHtml(strStripHtml != _T("") ? true : false)
			.SetImages(strImages != _T("") ? true : false)
			.SetReadOnly(strReadOnly != _T("") ? true : false)
			.SetDragOps(XHTMLTREE_DO_DEFAULT)
			.SetSeparatorColor(colorSeparator).SetBkColor(colorBK);
		//.SetDropCursors(IDC_NODROP, IDC_DROPCOPY, IDC_DROPMOVE);
		if (strImages != _T(""))
		{
			CComBSTR bstrImageURL(L"");
			pWindowNode->get_Attribute(CComBSTR(L"ImageURL"), &bstrImageURL);
			pTangramDesignView->m_pXHtmlTree->m_strImageURL = OLE2T(bstrImageURL);
			pWindowNode->get_Attribute(CComBSTR(L"ImageTarget"), &bstrImageURL);
			CString strImage = OLE2T(bstrImageURL);
			if (strImage != _T(""))
			{
				//
				CString strPath = theApp.m_strTempPath;
				strPath += _T("TangramTreeNode");// OLE2T(bstrImageURL);
				strPath += strImage;
				pTangramDesignView->m_pXHtmlTree->m_strImageTarget = strPath;
				int nPos = strPath.ReverseFind('\\');
				CString strDir = strPath.Left(nPos);
				SHCreateDirectory(NULL, strDir);
				if (::PathFileExists(strPath) == false)
				{
					//CComPtr<ITangramRestObj> p;
					//p.CoCreateInstance(L"TangramEx.TangramRestObj.1");
					//if (p)
					//{
					//	ITangramRestObj* m_pTangramRestObj = p.p;
					//	m_pTangramRestObj->AddRef();
					//	m_pTangramRestObj->put_CloudAddinRestNotify((IRestNotify*)pTangramDesignView->m_pWndNode);
					//	m_pTangramRestObj->UploadFile(false, CComBSTR(pTangramDesignView->m_pXHtmlTree->m_strImageURL), CComBSTR(L""), CComBSTR(strPath));
					//}
				}
				else
				{
					CImage image;
					image.Load(strPath);
					int nWidet = image.GetWidth();
					int nHeight = image.GetHeight();

					COLORREF color = image.GetTransparentColor();

					HBITMAP hBitmap;
					CBitmap* pBitmap;
					hBitmap = (HBITMAP)::LoadImage(::AfxGetInstanceHandle(), (LPCTSTR)strPath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
					pBitmap = new CBitmap;
					pBitmap->Attach(hBitmap);
					pTangramDesignView->m_pXHtmlTree->m_Images.Add(pBitmap, color);
					delete pBitmap;
				}
			}

			//AfxSetResourceHandle(theApp.m_hInstance);
			////HBITMAP hFontSelect = (HBITMAP)LoadImage(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDB_FONTSELECT), IMAGE_BITMAP, 0, 0, 0);			
			///*
			//CImageList* m_pImageList = new CImageList();

			//HBITMAP hBitmap;
			//CBitmap* pBitmap;
			//CString strFileName;
			//strFileName.Format("%s//res//root.bmp", szPath);
			//hBitmap=(HBITMAP)::LoadImage(::AfxGetInstanceHandle(),       (LPCTSTR)strFileName,IMAGE_BITMAP,0,0,LR_LOADFROMFILE);
			//pBitmap = new CBitmap;
			//pBitmap->Attach(hBitmap);
			//m_pImageList->Add(pBitmap,RGB(0,0,0));
			//delete pBitmap;
			//*/
			////AFX_MANAGE_STATE(AfxGetStaticModuleState());
			////pTangramDesignView->m_pXHtmlTree->SetBKImage(hFontSelect);
			//if (pTangramDesignView->m_pXHtmlTree->m_Images.m_hImageList == NULL)
			//{
			//	//pTangramDesignView->m_pXHtmlTree->m_Images.Create(IDB_TREENODEBITMAP, 16, 9, RGB(0, 0, 0));
			//}
			//pTangramDesignView->m_pXHtmlTree->SetImageList(&pTangramDesignView->m_pXHtmlTree->m_Images, TVSIL_NORMAL);
		}

		CTangramXmlParse* pParse = pWindowNode->m_pHostParse;
		int nCol = pParse->GetCount();
		if (nCol)
		{
			CTangramXmlParse* pTreeNodeParse = pParse->GetChild(_T("TreeNode"));
			if (pTreeNodeParse)
			{
				CComBSTR bstrTag(L"");
				pWindowNode->get_Attribute(CComBSTR(L"doctag"), &bstrTag);
				CString strTag = OLE2T(bstrTag);
				if (strTag.CompareNoCase(_T("tangramdesigner")) == 0)
				{
					CString strKey = L"currentdesignxml";
					auto it = m_mapValInfo.find(strKey);
					if (it != m_mapValInfo.end())
					{
						CString strXml = OLE2T(it->second.bstrVal);
						::VariantClear(&it->second);
						m_mapValInfo.erase(it);
						pWindowNode->m_pRootObj->m_pDocTreeCtrl = pTangramDesignView->m_pXHtmlTree;
						CTangramXmlParse* pParse = new CTangramXmlParse();
						if (pParse->LoadXml(strXml))
						{
							pTangramDesignView->m_pXHtmlTree->m_pHostXmlParse = pParse;
							pWindowNode->m_pRootObj->m_pDocXmlParseNode = pParse;
							pTangramDesignView->m_pXHtmlTree->m_hFirstRoot = pTangramDesignView->m_pXHtmlTree->LoadXmlFromXmlParse(pParse);
						}
					}
				}
				else
				{
					CComBSTR bstrType(L"");
					pWindowNode->get_Attribute(CComBSTR(L"doctype"), &bstrType);
					CString strType = OLE2T(bstrType);
					if (bstrType != _T(""))
					{
						CXHtmlDraw::XHTMLDRAW_APP_COMMAND AppCommands[] =
						{
							{ pTangramDesignView->m_pXHtmlTree->m_hWnd, WM_TANGRAMDESIGNERCMD, 1992, _T("WM_TANGRAM_DESIGNERCMD") },
							{ pTangramDesignView->m_pXHtmlTree->m_hWnd, WM_TANGRAMDESIGNERCMD, 1963, _T("WM_TANGRAM_DESIGNERCMD2") },
						};

						pTangramDesignView->m_pXHtmlTree->m_Links.SetAppCommands(AppCommands, sizeof(AppCommands) / sizeof(AppCommands[0]));

						if(m_pDocDOMTree==nullptr)
							m_pDocDOMTree = pTangramDesignView->m_pXHtmlTree;
					}
					else
					{
						CString strXml = pTreeNodeParse->xml();
						CTangramXmlParse* pParse = new CTangramXmlParse();
						pParse->LoadXml(strXml);
						pTangramDesignView->m_pXHtmlTree->m_hFirstRoot = pTangramDesignView->m_pXHtmlTree->LoadXmlFromXmlParse(pParse);
					}
				}
			}

			//pTreeNodeParse = pParse->GetChild(_T("TreeNodePlugin"));
			//if (pTreeNodeParse)
			//{
			//	CComBSTR bstrXML(L"");
			//	pTreeNodeParse->get_xml(&bstrXML);
			//	TElem subItemClick = subItem2.subnode(_T("ItemClick"));
			//	TElem subItemDoubleClick = subItem2.subnode(_T("ItemDoubleClick"));
			//	//pTangramDesignView->m_pXHtmlTree->m_hFirstRoot = pTangramDesignView->m_pXHtmlTree->LoadXmlFromString(OLE2T(bstrXML), CTangramHtmlTreeWnd::ConvertToUnicode);
			//}
		}
	}
	bRet = true;

	//Very important:
	if (pTangramDesignView&&::IsWindow(pTangramDesignView->m_hWnd))
		pTangramDesignView->SendMessage(WM_INITIALUPDATE);

	////////////////////////////////////////////////////////////////////////////////////////////////

	pWindowNode->m_pRootObj->m_mapLayoutNodes[pWindowNode->m_strName] = pWindowNode;
	if (pWindowNode&&pWindowNode->m_strID.CompareNoCase(_T("treeview")))
	{
		int nCol = pWindowNode->m_pHostParse->GetCount();

		pWindowNode->m_nRows = 1;
		pWindowNode->m_nCols = nCol;

		if (nCol)
		{
			pWindowNode->m_nViewType = TabbedWnd;
			if (pWindowNode->m_nActivePage<0 || pWindowNode->m_nActivePage>nCol - 1)
				pWindowNode->m_nActivePage = 0;
			CWnd* pView = nullptr;
			CWndNode* pObj = nullptr;
			int j = 0;
			for (int i = 0; i < nCol; i++)
			{
				CTangramXmlParse* pChild = pWindowNode->m_pHostParse->GetChild(i);
				CString _strName = pChild->name();
				CString strName = pChild->attr(TGM_NAME, _T(""));
				if (_strName.CompareNoCase(_T("node")) == 0)
				{
					strName.Trim();
					strName.MakeLower();

					pObj = new CComObject<CWndNode>;
					if (pWindowNode->m_pFrame)
					{
						pObj->m_pRootObj = pWindowNode->m_pRootObj;
						pObj->m_pFrame = pWindowNode->m_pFrame;
					}
					pObj->m_pHostParse = pChild;
					InitWndNode(pObj);

					pWindowNode->AddChildNode(pObj);

					pObj->m_nCol = j;

					if (pObj->m_pObjClsInfo)
					{
						pContext->m_pNewViewClass = pObj->m_pObjClsInfo;
						pView = (CWnd*)pContext->m_pNewViewClass->CreateObject();
						pView->Create(NULL, _T(""), WS_VISIBLE | WS_CHILD, rect, pTangramDesignView, 0, pContext);
						HWND m_hChild = (HWND)::SendMessage(pTangramDesignView->m_hWnd, WM_CREATETABPAGE, (WPARAM)pView->m_hWnd, (LPARAM)LPCTSTR(pObj->m_strCaption));
					}
					j++;
				}
			}
			::SendMessage(pTangramDesignView->m_hWnd, WM_ACTIVETABPAGE, (WPARAM)pWindowNode->m_nActivePage, (LPARAM)1);
		}
	}

	pWindowNode->m_pHostWnd->SetWindowText(pWindowNode->m_strWebObjName);

	if (pWindowNode->m_nViewType == TabbedWnd&&pWindowNode->m_pParentObj&&pWindowNode->m_pParentObj->m_nViewType == Splitter)
	{
		if (pWindowNode->m_pHostWnd)
			pWindowNode->m_pHostWnd->ModifyStyleEx(WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE, 0);
	}
	if (m_pWndNode->m_pPage)
		m_pWndNode->m_pPage->Fire_NodeCreated(pWindowNode);

	return bRet;
}

LRESULT CALLBACK CTangramApp::TangramWndProc(_In_ HWND hWnd,UINT msg, _In_ WPARAM wParam,_In_ LPARAM lParam)
{
	switch (msg)
	{
		case WM_TANGRAMUCMAMSG:
			{
				UCMAMSGInfo* pUCMAMSGInfo = (UCMAMSGInfo*)wParam;
				if (lParam == 0)
				{
					CString strMsg = pUCMAMSGInfo->m_strMsg;
					strMsg.Replace(_T("tangramCmd:"),_T(""));
					CString strSip = pUCMAMSGInfo->m_strData;
					CString s = strSip;
					pUCMAMSGInfo = new UCMAMSGInfo();
					pUCMAMSGInfo->m_strMsg = strMsg;
					pUCMAMSGInfo->m_strData = strSip;
					::PostMessage(theApp.m_hActiveWnd, WM_TANGRAMUCMAMSG, (WPARAM)pUCMAMSGInfo, 0);
				}
				//else
				//	delete pUCMAMSGInfo;
			}
			return 0;
		case WM_CLOSE:
			{
				::ShowWindow(theApp.m_pHostCore->m_hHostWnd, SW_HIDE);
			}
			return 0;
		case WM_DESTROY:
			{
				if (hWnd==theApp.m_pHostCore->m_hHostWnd&&theApp.m_pDesignerFrame)
				{
					HWND hWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), _T(""), WS_CHILD, 0, 0, 0, 0, theApp.m_pHostCore->m_hHostWnd, NULL, theApp.m_hInstance, NULL);
					HWND hChildWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), _T(""), WS_CHILD, 0, 0, 0, 0, hWnd, NULL, theApp.m_hInstance, NULL);
					if (hWnd&&hChildWnd)
					{
						theApp.m_pDesignerFrame->ModifyHost((long)hChildWnd);
						if (::DestroyWindow(hWnd))
						{
							theApp.m_pDesignerFrame = nullptr;
							theApp.m_pDesignerWndPage = nullptr;
						}
					}
				}
			}
			break;
		case WM_WINDOWPOSCHANGED:
			if (theApp.m_pHostCore->m_hChildHostWnd)
			{
				RECT rc;
				::GetClientRect(theApp.m_pHostCore->m_hHostWnd, &rc);
				::SetWindowPos(theApp.m_pHostCore->m_hChildHostWnd, NULL, 0, 0, rc.right, rc.bottom, SWP_NOACTIVATE | SWP_NOREDRAW);
			}
			break;
		case WM_OPENDOCUMENT: 
			{
				theApp.m_pHostCore->OnOpenDoc(wParam);
			}
			break;
		case WM_OFFICEOBJECTCREATED:
			{
				HWND hWnd = (HWND)wParam;
				((OfficeCloudPlus::CCloudAddin*)theApp.m_pHostCore)->ConnectOfficeObj(hWnd);
			}
			break;
		case WM_TANGRAMECLIPSEINFO:
			{
				//InitEclipse();
			}
			break;
		case WM_TANGRAMMSG:
			if (wParam == 1963 && lParam == 1222)
			{
				OfficeCloudPlus::OutLookPlus::COutLookCloudAddin* pAddin = (OfficeCloudPlus::OutLookPlus::COutLookCloudAddin*)theApp.m_pHostCore;
				if (pAddin->m_pActiveOutlookExplorer)
				{
					pAddin->m_pActiveOutlookExplorer->SetDesignState();
				}
			}
			else
			{
				HWND hwnd = (HWND)wParam;
				if (::IsWindow(hwnd))
					::DestroyWindow(hwnd);
			}
			break;
		case WM_TANGRAMNEWOUTLOOKOBJ:
			{
				using namespace OfficeCloudPlus::OutLookPlus;
				int nType = wParam;
				HWND hWnd = ::GetActiveWindow();
				if (nType)
				{
					COutLookExplorer* pOutLookPlusItemWindow = (COutLookExplorer*)lParam;
					COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
					pOutLookPlusItemWindow->m_strKey = pAddin->m_strCurrentKey;
					pAddin->m_mapOutLookPlusExplorerMap[hWnd] = pOutLookPlusItemWindow;
					pOutLookPlusItemWindow->m_hWnd = hWnd;
				}
			}
			break;
		case WM_TANGRAMACTIVEINSPECTORPAGE:
			{
				using namespace OfficeCloudPlus::OutLookPlus;
				COutLookInspector* pOutLookPlusItemWindow = (COutLookInspector*)wParam;
				pOutLookPlusItemWindow->ActivePage();
			}
			break;
		case WM_TANGRAMITEMLOAD:
			{
				using namespace OfficeCloudPlus::OutLookPlus;
				COutLookCloudAddin* pAddin = (COutLookCloudAddin*)theApp.m_pHostCore;
				HWND hWnd = ::GetActiveWindow();
				auto it = pAddin->m_mapOutLookPlusExplorerMap.find(hWnd);
				if (it != pAddin->m_mapOutLookPlusExplorerMap.end())
				{
					COutLookExplorer* pExplorer = it->second;
					if (pExplorer->m_pInspectorContainerWnd == nullptr)
					{
						HWND _hWnd = ::FindWindowEx(hWnd, NULL, _T("rctrl_renwnd32"), NULL);
						if (_hWnd)
						{
							_hWnd = ::FindWindowEx(_hWnd, NULL, _T("AfxWndW"), NULL);
							if (_hWnd)
							{
								pExplorer->m_pInspectorContainerWnd = new CInspectorContainerWnd();
								pExplorer->m_pInspectorContainerWnd->SubclassWindow(_hWnd);
							}
						}
					}

					long nKey = wParam;
					auto it = pAddin->m_mapTangramInspectorItem.find(nKey);
					if (it != pAddin->m_mapTangramInspectorItem.end())
					{
						CInspectorItem* pItem = (CInspectorItem*)wParam;
						if (pExplorer->m_pInspectorContainerWnd)
						{
							pExplorer->m_pInspectorContainerWnd->m_strXml = pItem->m_strXml;
							::PostMessage(pExplorer->m_pInspectorContainerWnd->m_hWnd, WM_TANGRAMITEMLOAD, 0, 0);
						}
					}
				}
			}
			break;
		default:
			break;
	}
	return ::DefWindowProc(hWnd,msg, wParam,lParam);
}

LRESULT CTangramApp::CBTProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LRESULT hr = CallNextHookEx(theApp.m_hCBTHook, nCode, wParam, lParam);
	HWND hWnd = (HWND)wParam;
	switch (nCode)
	{
	case HCBT_CREATEWND:
	{
		CBT_CREATEWND* pCreateWnd = (CBT_CREATEWND*)lParam;

		LPCTSTR lpszName = pCreateWnd->lpcs->lpszName;
		::GetClassName(hWnd, theApp.m_szBuffer, MAX_PATH);
		CString strClassName = CString(theApp.m_szBuffer);

		//if (pCreateWnd->lpcs->hwndParent == NULL&&theApp.m_strSWTClassName != _T("") && strClassName == theApp.m_strSWTClassName)
		//{
		//	EclipseCloudPlus::EclipsePlus::CEclipseCloudAddin* pAddin = ((EclipseCloudPlus::EclipsePlus::CEclipseCloudAddin*)theApp.m_pHostCore);
		//	if (pAddin->m_bClose == false)
		//	{
		//		::PostMessage(pAddin->m_hHostWnd, WM_ECLIPSEMAINWNDCREATED, (WPARAM)hWnd, 0);
		//	}			
		//	break;
		//}
		
		if (theApp.m_bOfficeApp&&HIWORD(pCreateWnd->lpcs->lpszClass)&&theApp.m_pHostCore)
		{
			((CCloudAddin*)theApp.m_pHostCore)->WindowCreated(strClassName, lpszName, pCreateWnd->lpcs->hwndParent, hWnd);
		}
		else if(theApp.m_pTangramProxy)
			theApp.m_pTangramProxy->WindowCreated(strClassName, lpszName, pCreateWnd->lpcs->hwndParent, hWnd);
	}
	break;
	case HCBT_DESTROYWND:
	{
		if (theApp.m_bOfficeApp&&theApp.m_pHostCore)
		{
			((CCloudAddin*)theApp.m_pHostCore)->WindowDestroy(hWnd);
		}
		else if (theApp.m_pTangramProxy)
			theApp.m_pTangramProxy->WindowDestroy(hWnd);

		auto it = theApp.m_mapWindowPage.find(hWnd);
		if (it != theApp.m_mapWindowPage.end())
		{
			CWndPage* pPage = it->second;
			if (pPage->m_mapFrame.size() == 0)
			{
				theApp.m_mapWindowPage.erase(it);
				delete pPage;
			}
		}
	}
	break;
	case HCBT_MINMAX:
		switch (lParam)
		{
			case SW_MINIMIZE:
			{
				if (::GetWindowLong(hWnd, GWL_EXSTYLE)&WS_EX_MDICHILD)
					::PostMessage(hWnd, WM_MDICHILDMIN, 0, 0);
			}
			break;
		}
		break;
	case HCBT_SETFOCUS:
		{
			if(theApp.m_bOfficeApp)
				((CCloudAddin*)theApp.m_pHostCore)->SetFocus(hWnd);
		}
		break;
	case HCBT_ACTIVATE:
		{
			if (theApp.m_bOfficeApp)
			{
				if (((CCloudAddin*)theApp.m_pHostCore)->OnActiveOfficeObj(hWnd))
					break;
			}

			if (theApp.m_bClose == false)
			{
				theApp.m_hActiveWnd = hWnd;
			}
		}
		break;
	}
	return hr;
}

LRESULT CALLBACK CTangramApp::GetMessageProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LPMSG lpMsg = (LPMSG)lParam;
	DWORD dwID = ::GetCurrentThreadId();
	TangramThreadInfo* pThreadInfo = theApp.GetThreadInfo(dwID);
	//if ((nCode >= 0) && PM_NOREMOVE == wParam)
	//{
	//	switch (lpMsg->message)
	//	{
	//	case WM_QUIT:
	//		{
	//			ATLTRACE(_T("WM_ENABLE\n"));
	//		}
	//		break;
	//	}
	//}
	if ((nCode >= 0) && PM_REMOVE == wParam)
	{
		switch (lpMsg->message)
		{
		case WM_KEYDOWN:
		{
			switch (lpMsg->wParam)
			{
			case VK_TAB:
				if (theApp.m_bWinFormActived&&theApp.m_bEnableProcessFormTabKey)
				{
					break;
				}
				if (theApp.m_pWndFrame&&theApp.m_pWndNode)
				{
					CNodeWnd* pWnd = (CNodeWnd*)theApp.m_pWndNode->m_pHostWnd;
					if (pWnd&&!::IsChild(pWnd->m_hWnd, lpMsg->hwnd))
					{
						if (pWnd->m_bBKWnd)
						{
							if (!::IsChild(theApp.m_pMDIClientBKWnd->m_hWnd, lpMsg->hwnd))
							{
								theApp.m_pWndNode = nullptr;
								theApp.m_pWndFrame = nullptr;
							}
							else if (pWnd->PreTranslateMessage(lpMsg))
								lpMsg->hwnd = NULL;
						}
						else
						{
							theApp.m_pWndNode = nullptr;
							theApp.m_pWndFrame = nullptr;
						}
					}
					else if (pWnd&&pWnd->PreTranslateMessage(lpMsg))
					{
						lpMsg->hwnd = NULL;
						lpMsg->lParam = 0;
						lpMsg->wParam = 0;
						lpMsg->message = 0;
					}
				}
				break;
			case VK_PRIOR:
			case VK_NEXT:
			case VK_HOME:
			case VK_END:
			case VK_LEFT:
			case VK_UP:
			case VK_RIGHT:
			case VK_DOWN:
				if (theApp.m_pWndFrame&&theApp.m_pWndNode)
				{
					CNodeWnd* pWnd = (CNodeWnd*)theApp.m_pWndNode->m_pHostWnd;
					if (pWnd->m_bBKWnd)
					{
						if (theApp.m_pMDIClientBKWnd&&::IsChild(theApp.m_pMDIClientBKWnd->m_hWnd, lpMsg->hwnd) && pWnd->PreTranslateMessage(lpMsg))
						{
							lpMsg->hwnd = NULL;
							lpMsg->wParam = 0;
						}
						else
						{
							theApp.m_pWndNode = nullptr;
							theApp.m_pWndFrame = nullptr;
							DispatchMessage(lpMsg);
							lpMsg->hwnd = NULL;
							lpMsg->wParam = 0;
						}
						break;
					}
					if (pWnd&&pWnd->PreTranslateMessage(lpMsg))
					{
						lpMsg->hwnd = NULL;
						lpMsg->lParam = 0;
						lpMsg->wParam = 0;
						lpMsg->message = 0;
					}
				}
				break;
			case VK_DELETE:
				if (theApp.m_pWndNode)
				{
					CNodeWnd* pWnd = (CNodeWnd*)theApp.m_pWndNode->m_pHostWnd;
					HWND hWnd = pWnd->m_hWnd;
					if (pWnd->m_bBKWnd)
					{
						if (theApp.m_pMDIClientBKWnd&&::IsChild(theApp.m_pMDIClientBKWnd->m_hWnd, lpMsg->hwnd))
							pWnd->PreTranslateMessage(lpMsg);
						else
							DispatchMessage(lpMsg);
						lpMsg->hwnd = NULL;
						lpMsg->wParam = 0;
						break;
					}
					if (::IsChild(hWnd, lpMsg->hwnd) == false)
					{
						DispatchMessage(lpMsg);
						lpMsg->hwnd = NULL;
						lpMsg->wParam = 0;
						break;
					}
					if (theApp.m_pWndNode->m_nViewType == ActiveX)
					{
						pWnd->PreTranslateMessage(lpMsg);
						lpMsg->hwnd = NULL;
						lpMsg->wParam = 0;
						break;
					}
					DispatchMessage(lpMsg);
					lpMsg->hwnd = NULL;
					lpMsg->wParam = 0;
				}

				break;
			case VK_BACK:
			{
				if (theApp.m_pWndNode)
				{
					CNodeWnd* pWnd = (CNodeWnd*)theApp.m_pWndNode->m_pHostWnd;
					if (pWnd)
					{
						HWND h = pWnd->m_hWnd;
						if (::IsChild(h, lpMsg->hwnd))
						{
							if (pWnd->PreTranslateMessage(lpMsg))
							{
								lpMsg->hwnd = NULL;
								lpMsg->lParam = 0;
								lpMsg->wParam = 0;
								lpMsg->message = 0;
							}
						}
						else
						{
							theApp.m_pWndNode = NULL;
							TranslateMessage(lpMsg);
						}
					}
				}
			}
			break;
			case VK_RETURN:
			{
				if (theApp.m_pWndFrame&&theApp.m_pWndNode)
				{
					CNodeWnd* pWnd = (CNodeWnd*)theApp.m_pWndNode->m_pHostWnd;
					if (pWnd&&::IsChild(pWnd->m_hWnd, lpMsg->hwnd) == false)
					{
						theApp.m_pWndNode = nullptr;
						theApp.m_pWndFrame = nullptr;
					}
					else if (pWnd)
					{
						TranslateMessage(lpMsg);
						lpMsg->hwnd = NULL;
						lpMsg->lParam = 0;
						lpMsg->wParam = 0;
						lpMsg->message = 0;
						break;
					}
				}
				TranslateMessage(lpMsg);
				if (theApp.m_strExeName != _T("devenv"))
				{
					DispatchMessage(lpMsg);
					lpMsg->hwnd = NULL;
					lpMsg->lParam = 0;
					lpMsg->wParam = 0;
					lpMsg->message = 0;
				}
			}
			break;
			case 0x41://Ctrl+A
			case 0x43://Ctrl+C
			case 0x56://Ctrl+V
			case 0x58://Ctrl+X
			case 0x5a://Ctrl+Z
				if (::GetKeyState(VK_CONTROL) < 0)  // control pressed
				{
					if (theApp.m_pWndNode&&theApp.m_pWndNode->m_pHostWnd&&!::IsWindow(theApp.m_pWndNode->m_pHostWnd->m_hWnd))
					{
						theApp.m_pWndNode = nullptr;
					}
					if (theApp.m_pWndNode && (theApp.m_pWndNode->m_nViewType == ActiveX || theApp.m_pWndNode->m_strID.CompareNoCase(_T("hostview")) == 0))
					{
						CNodeWnd* pWnd = (CNodeWnd*)theApp.m_pWndNode->m_pHostWnd;
						HWND hWnd = pWnd->m_hWnd;
						if (pWnd->m_bBKWnd&&theApp.m_pMDIClientBKWnd)
							hWnd = theApp.m_pMDIClientBKWnd->m_hWnd;
						if (::IsChild(hWnd, lpMsg->hwnd))
						{
							pWnd->PreTranslateMessage(lpMsg);
							lpMsg->hwnd = NULL;
							lpMsg->wParam = 0;
							break;
						}
					}
					if (theApp.m_pWndNode)
					{
						TranslateMessage(lpMsg);
						lpMsg->wParam = 0;
					}
				}
				break;
			}
		}
		break;
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
		case WM_SETWNDFOCUSE:
			theApp.m_pHostCore->ProcessMsg(lpMsg);
			break;
		case WM_MDICHILDMIN:
			::BringWindowToTop(lpMsg->hwnd);
			break;
		case WM_DOWNLOAD_MSG:
		{
			CloudUtilities::CDownLoadObj* pObj = (CloudUtilities::CDownLoadObj*)lpMsg->wParam;
			if (pObj)
			{
				delete pObj;
			}
		}
		break;
		case WM_NAVIXTML:
		{
			RECT rc;
			HWND hWnd = ::GetParent(lpMsg->hwnd);
			::GetClientRect(hWnd, &rc);
			::SetWindowPos(lpMsg->hwnd, HWND_BOTTOM, rc.left, rc.top, rc.right + 1, rc.bottom, SWP_NOZORDER | SWP_FRAMECHANGED);
			::SetWindowPos(lpMsg->hwnd, HWND_BOTTOM, rc.left, rc.top, rc.right, rc.bottom, SWP_NOZORDER | SWP_FRAMECHANGED);
		}
		break;
		case WM_TANGRAM_WEBNODEDOCCOMPLETE:
		{
			auto it = theApp.m_mapWindowPage.find(lpMsg->hwnd);
			if (it != theApp.m_mapWindowPage.end())
				it->second->OnNodeDocComplete(lpMsg->wParam);
		}
		break;
#ifndef _WIN64
		case WM_INITPROPERTYGRID:
		{
			VSCloudPlus::VisualStudioPlus::CVSCloudAddin* pAddin = (VSCloudPlus::VisualStudioPlus::CVSCloudAddin*)theApp.m_pHostCore;
			pAddin->InitPropertyGridCtrl(lpMsg->hwnd);
		}
#endif
		break;		
		}
	}
	return CallNextHookEx(pThreadInfo->m_hGetMessageHook, nCode, wParam, lParam);
}

CString CTangramApp::EncodeFileToBase64(CString strSRC)
{
	DWORD dwDesiredAccess = GENERIC_READ;
	DWORD dwShareMode = FILE_SHARE_READ;
	DWORD dwFlagsAndAttributes = FILE_ATTRIBUTE_NORMAL;
	HANDLE hFile = ::CreateFile(strSRC, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_FLAG_RANDOM_ACCESS, NULL);

	if (hFile == INVALID_HANDLE_VALUE)
	{
		TRACE(_T("ERROR: CreateFile failed - %s\n"), strSRC);
		return _T("");
	}
	else
	{
		DWORD dwFileSizeHigh = 0;
		__int64 qwFileSize = GetFileSize(hFile, &dwFileSizeHigh);
		qwFileSize |= (((__int64)dwFileSizeHigh) << 32);
		DWORD dwFileSize = qwFileSize;
		if ((dwFileSize == 0) || (dwFileSize == INVALID_FILE_SIZE))
		{
			TRACE(_T("ERROR: GetFileSize failed - %s\n"), strSRC);
			return _T("");
		}
		else
		{
			PBYTE buffer = new BYTE[dwFileSize];
			memset(buffer, 0, (dwFileSize)*sizeof(BYTE));
			DWORD dwBytesRead = 0;
			if (!ReadFile(hFile, buffer, dwFileSize, &dwBytesRead, NULL))
			{
				TRACE(_T("ERROR: ReadFile failed - %s\n"), strSRC);
			}
			else
			{
				int nMaxLineLen = 256;
				char *pDstInfo = new char[dwFileSize * 2];
				memset(pDstInfo, 0, dwFileSize * 2);
				Base64Encode(buffer, pDstInfo, dwFileSize, nMaxLineLen);
				CString strInfo = CA2W(pDstInfo);
				delete[] pDstInfo;

				delete[] buffer;
				return strInfo;
			}
		}
	}

	return _T("");
}

CString CTangramApp::Encode(CString strSRC, BOOL bEnCode)
{
	if (bEnCode)
	{
		LPCWSTR srcInfo = strSRC;
		std::string strSrc = (LPCSTR)CW2A(srcInfo, CP_UTF8);
		int nSrcLen = strSrc.length();
		int nDstLen = Base64EncodeGetRequiredLength(nSrcLen);
		char *pDstInfo = new char[nSrcLen * 2];
		memset(pDstInfo, 0, nSrcLen * 2);
		ATL::Base64Encode((BYTE*)strSrc.c_str(), nSrcLen, pDstInfo, &nDstLen);
		CString strInfo = CA2W(pDstInfo);
		delete[] pDstInfo;
		return strInfo;
	}
	else
	{
		long nSrcSize = strSRC.GetLength();
		BYTE *pDecodeStr = new BYTE[nSrcSize];
		memset(pDecodeStr, 0, nSrcSize);
		int nLen = nSrcSize;
		ATL::Base64Decode(CW2A(strSRC), nSrcSize, pDecodeStr, &nLen);
		CString str = CA2W((char*)pDecodeStr, CP_UTF8);
		delete[] pDecodeStr;
		pDecodeStr = NULL;
		return str;
	}
}

void CTangramApp::SetHook(DWORD ThreadID)
{
	TangramThreadInfo* pThreadInfo = GetThreadInfo(ThreadID);
	if (pThreadInfo&&pThreadInfo->m_hGetMessageHook == NULL)
	{
		pThreadInfo->m_hGetMessageHook = SetWindowsHookEx(WH_GETMESSAGE, GetMessageProc, NULL, ThreadID);
	}
}

void CTangramApp::InitWndNode(CWndNode* pObj)
{
	if (pObj)
	{
		pObj->m_nHeigh = pObj->m_pHostParse->attrInt(TGM_HEIGHT, 0);
		pObj->m_nWidth = pObj->m_pHostParse->attrInt(TGM_WIDTH, 0);

		pObj->m_nStyle = pObj->m_pHostParse->attrInt(TGM_STYLE, 0);
		pObj->m_nActivePage = pObj->m_pHostParse->attrInt(TGM_ACTIVE_PAGE, 0);
		pObj->m_strCaption = pObj->m_pHostParse->attr(TGM_CAPTION, _T(""));
		pObj->m_strName = pObj->m_pHostParse->attr(TGM_NAME, _T(""));
		if (pObj->m_strName == _T(""))
		{
			CString strName = _T("");
			strName.Format(_T("Node_%p"), (LONGLONG)pObj);
			pObj->m_strName = strName;
		}
		if (pObj->m_pFrame&&pObj->m_pFrame->m_pPage)
		{
			pObj->m_pPage = pObj->m_pFrame->m_pPage;
			pObj->m_strWebObjName = pObj->m_pFrame->m_strFrameName + _T("_") + theApp.m_strCurrentKey + _T("_") + pObj->m_strName;
			auto it2 = pObj->m_pPage->m_mapNode.find(pObj->m_strWebObjName);
			if (it2 == pObj->m_pPage->m_mapNode.end())
			{
				pObj->m_pPage->m_mapNode[pObj->m_strWebObjName] = pObj;
			}
		}
		pObj->m_strID = pObj->m_pHostParse->attr(TGM_NODE_TYPE, _T(""));
		pObj->m_strID.MakeLower();
		pObj->m_strID.Trim();
		pObj->m_strCnnID = pObj->m_pHostParse->attr(TGM_CNN_ID, _T(""));
		pObj->m_strCnnID.MakeLower();
		pObj->m_strCnnID.Trim();
		TangramWndClsInfo* pWndClsInfo = nullptr;
		if (m_TabWndClassInfoDictionary.Lookup(LPCTSTR(pObj->m_strID), (void *&)pWndClsInfo))
			pObj->m_pObjClsInfo = pWndClsInfo->m_pTabWndClsInfo;
	}

	if (pObj->m_pObjClsInfo == nullptr)
		pObj->m_pObjClsInfo = RUNTIME_CLASS(CNodeWnd);
}

CString CTangramApp::GetPropertyFromObject(IDispatch* pObj, CString strPropertyName)
{
	CString strRet = _T("");
	if (pObj)
	{
		BSTR szMember = strPropertyName.AllocSysString();
		DISPID dispid = -1;
		HRESULT hr = pObj->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_USER_DEFAULT, &dispid);
		if (hr == S_OK)
		{
			DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
			VARIANT result = { 0 };
			EXCEPINFO excepInfo;
			memset(&excepInfo, 0, sizeof excepInfo);
			UINT nArgErr = (UINT)-1;
			HRESULT hr = pObj->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &result, &excepInfo, &nArgErr);
			if (S_OK == hr && VT_BSTR == result.vt)
			{
				strRet = OLE2T(result.bstrVal);
			}
			::VariantClear(&result);
		}
	}
	return strRet;
}

CString CTangramApp::GetXmlData(CString strName, CString strXml)
{
	if (strName == _T("")||strXml==_T(""))
		return _T("");
	int nLength = strName.GetLength();
	CString strKey = _T("<") + strName + _T(">");
	int nPos = strXml.Find(strKey);
	if (nPos != -1)
	{
		CString strData1 = strXml.Mid(nPos);
		strKey = _T("</") + strName + _T(">");
		nPos = strData1.Find(strKey);
		if (nPos != -1)
			return strData1.Left(nPos + nLength + 3);
	}
	return _T("");
}

CString CTangramApp::GetNewGUID()
{
	GUID   m_guid;
	CString   strGUID = _T("");
	if (S_OK == ::CoCreateGuid(&m_guid))
	{
		strGUID.Format(_T("%08X-%04X-%04x-%02X%02X-%02X%02X%02X%02X%02X%02X"),
			m_guid.Data1, m_guid.Data2, m_guid.Data3,
			m_guid.Data4[0], m_guid.Data4[1],
			m_guid.Data4[2], m_guid.Data4[3],
			m_guid.Data4[4], m_guid.Data4[5],
			m_guid.Data4[6], m_guid.Data4[7]);
	}

	return strGUID;
}

CString CTangramApp::GetNames(IWndNode* _pNode)
{
	CString strRet = _T("");
	CWndNode* pNode = (CWndNode*)_pNode;
	if (pNode)
	{
		pNode = pNode->m_pRootObj;
		strRet += pNode->m_strName;
		strRet += _T(",");
		strRet += _GetNames(pNode);
	}
	return strRet;
}

CString CTangramApp::_GetNames(CWndNode* pNode)
{
	CString strRet = _T("");
	if (pNode)
	{
		for (auto it : pNode->m_vChildNodes)
		{
			CWndNode* pChildNode = it;
			strRet += pChildNode->m_strName;
			strRet += _T(",");
			strRet += _GetNames(pChildNode);
		}
	}
	return strRet;
}

CString CTangramApp::GetFileVer()
{
	DWORD dwHandle, InfoSize;
	CString strVersion;

	struct LANGANDCODEPAGE
	{
		WORD wLanguage;
		WORD wCodePage;
	}*lpTranslate;
	unsigned int cbTranslate = 0;

	TCHAR cPath[MAX_PATH] = { 0 };
	::GetModuleFileName(NULL, cPath, MAX_PATH);
	InfoSize = GetFileVersionInfoSize(cPath, &dwHandle);


	char *InfoBuf = new char[InfoSize];

	GetFileVersionInfo(cPath, 0, InfoSize, InfoBuf);
	VerQueryValue(InfoBuf, TEXT("\\VarFileInfo\\Translation"), (LPVOID*)&lpTranslate, &cbTranslate);

	TCHAR SubBlock[300] = { 0 };

	//ProductVersion
	//FileVersion

	wsprintf(SubBlock, TEXT("\\StringFileInfo\\%04x%04x\\ProductVersion"), lpTranslate[0].wLanguage, lpTranslate[0].wCodePage);

	TCHAR *lpBuffer = NULL;
	unsigned int dwBytes = 0;
	VerQueryValue(InfoBuf, SubBlock, (void**)&lpBuffer, &dwBytes);
	if (lpBuffer != NULL)
	{
		strVersion.Format(_T("%s"), (TCHAR*)lpBuffer);
	}

	delete[] InfoBuf;
	return strVersion;
}

void CTangramApp::AddClsInfo(CString m_strObjID, int nType, CRuntimeClass* pClsInfo)
{
	m_strObjID.MakeLower();

	TRACE(_T("---------- %s\n"), m_strObjID.GetBuffer());

	TangramWndClsInfo* pTabWndClsInfo = new TangramWndClsInfo;
	pTabWndClsInfo->m_nType = (ViewType)nType;
	pTabWndClsInfo->m_pTabWndClsInfo = pClsInfo;
	m_TabWndClassInfoDictionary[m_strObjID] = pTabWndClsInfo;
}

void CTangramApp::HostPosChanged(CWndNode* pHostViewNode)
{
	if (pHostViewNode == nullptr)
		return;
	CWndFrame* pFrame = pHostViewNode->m_pFrame;
	if (pFrame)
	{
		HWND hwnd = pFrame->m_hWnd;
		CWndNode* pTopNode = pFrame->m_pWorkNode;
		CWndFrame* _pFrame = pFrame;
		while (pTopNode->m_pRootObj != pTopNode)
		{
			_pFrame = pTopNode->m_pRootObj->m_pFrame;
			hwnd = _pFrame->m_hWnd;
			pTopNode = _pFrame->m_pWorkNode;
		}
		if (::IsWindow(hwnd) == false)
			return;

		CRect rt1;
		CWnd* pWnd = _pFrame->m_pWorkNode->m_pHostWnd;
		pWnd->GetWindowRect(&rt1);
		CWindow parentwnd;
		parentwnd.Attach(hwnd);
		_pFrame->GetParent().ScreenToClient(&rt1);
		_pFrame->m_bNeedOuterChanged = false;
		HDWP dwh = BeginDeferWindowPos(1);
		dwh = ::DeferWindowPos(dwh, hwnd, HWND_TOP,
			rt1.left,
			rt1.top,
			rt1.Width(),
			rt1.Height(),
			SWP_FRAMECHANGED | SWP_NOACTIVATE
			);
		//dwh = ::DeferWindowPos(dwh, pFrame->m_hWnd, HWND_TOP,
		//	rt1.left,
		//	rt1.top,
		//	rt1.Width(),
		//	rt1.Height(),
		//	SWP_FRAMECHANGED | SWP_NOACTIVATE
		//	);
		EndDeferWindowPos(dwh);
	}
}

LRESULT CTangramApp::Close(void)
{
	HRESULT hr = S_OK;
	int cConnections = m_pHostCore->m_vec.GetSize();

	for (int iConnection = 0; iConnection < cConnections; iConnection++)
	{
		CComPtr<IUnknown> punkConnection = m_pHostCore->m_vec.GetAt(iConnection);

		IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

		if (pConnection)
		{
			CComVariant varResult;
			DISPPARAMS params = { NULL, NULL, 0, 0 };
			hr = pConnection->Invoke(2, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
		}
	}

	if (m_pCloudAddinCLRProxy)
	{
		m_pCloudAddinCLRProxy->CloseForms();
	}

	for (auto it : m_mViewFactory)
	{
		if (it.first.CompareNoCase(_T("HostApp")))
			it.second->Release();
	}
	m_mViewFactory.clear();
	return S_OK;
}

void CTangramApp::InitDesignerTreeCtrl(CString strXml)
{
	if (strXml != _T("")&& m_pDocDOMTree)
	{
		m_pDocDOMTree->m_pHostXmlParse = new CTangramXmlParse();
		m_pDocDOMTree->m_pHostXmlParse->LoadXml(strXml);
		m_pDocDOMTree->m_hFirstRoot = m_pDocDOMTree->LoadXmlFromXmlParse(m_pDocDOMTree->m_pHostXmlParse);
		m_pDocDOMTree->ExpandAll();
	}
}

void CTangramApp::InitEventDic()
{
	if (theApp.m_mapEventType.size() == 0)
	{
		m_mapEventType[_T("OnClick")] = TangramClick;
		m_mapEventType[_T("OnDoubleClick")] = TangramDoubleClick;
		m_mapEventType[_T("OEnter")] = TangramEnter;
		m_mapEventType[_T("OnLeave")] = TangramLeave;
		m_mapEventType[_T("OnEnabledChanged")] = TangramEnabledChanged;
		m_mapEventType[_T("OnLostFocus")] = TangramLostFocus;
		m_mapEventType[_T("OnGotFocus")] = TangramGotFocus;
		m_mapEventType[_T("OnKeyUp")] = TangramKeyUp;
		m_mapEventType[_T("OnKeyDown")] = TangramKeyDown;
		m_mapEventType[_T("OnKeyPress")] = TangramKeyPress;
		m_mapEventType[_T("OnMouseClick")] = TangramMouseClick;
		m_mapEventType[_T("OnMouseDoubleClick")] = TangramMouseDoubleClick;
		m_mapEventType[_T("OnMouseDown")] = TangramMouseDown;
		m_mapEventType[_T("OnMouseEnter")] = TangramMouseEnter;
		m_mapEventType[_T("OnMouseHover")] = TangramMouseHover;
		m_mapEventType[_T("OnMouseLeave")] = TangramMouseLeave;
		m_mapEventType[_T("OnMouseMove")] = TangramMouseMove;
		m_mapEventType[_T("OnMouseUp")] = TangramMouseUp;
		m_mapEventType[_T("OnouseWheel")] = TangramMouseWheel;
		m_mapEventType[_T("OnTextChanged")] = TangramTextChanged;
		m_mapEventType[_T("OnVisibleChanged")] = TangramVisibleChanged;
		m_mapEventType[_T("OnClientSizeChanged")] = TangramClientSizeChanged;
		m_mapEventType[_T("OnSizeChanged")] = TangramSizeChanged;
		m_mapEventType[_T("OnParentChanged")] = TangramParentChanged;
		m_mapEventType[_T("OnResize")] = TangramResize;
	}
}

void CTangramApp::InitJava(int nIndex)
{
	switch (nIndex)
	{
	case 0:
	case 1:
		{
			CComPtr<IDispatch> pDisp;
			pDisp.CoCreateInstance(L"Tangram.JavaProxy.1");
			CComQIPtr<IJavaProxy> pProxy(pDisp);
			if (pProxy)
			{
				theApp.m_pJavaProxy = pProxy.Detach();
				theApp.m_pJavaProxy->AddRef();
				if(nIndex==0)
					theApp.m_pJavaProxy->InitEclipse();
			}
		}
		break;
	case 100:
		break;
	case 111:
		break;
	default:
		break;
	}
}

int CTangramApp::LoadCLR()
{
	if (theApp.m_pCloudAddinCLRProxy == nullptr&&m_pClrHost == nullptr)
	{
		HMODULE	hMscoreeLib = LoadLibrary(TEXT("mscoree.dll"));
		if (hMscoreeLib)
		{
			HRESULT hrStart = 0;
			ICLRMetaHost* m_pMetaHost = NULL;
			TangramCLRCreateInstance CLRCreateInstance = (TangramCLRCreateInstance)GetProcAddress(hMscoreeLib, "CLRCreateInstance");
			if (CLRCreateInstance)
			{
				hrStart = CLRCreateInstance(CLSID_CLRMetaHost, IID_ICLRMetaHost, (LPVOID *)&m_pMetaHost);
				CString strVer = _T("v4.0.30319");
				ICLRRuntimeInfo * lpRuntimeInfo = nullptr;
				hrStart = m_pMetaHost->GetRuntime(strVer.AllocSysString(),IID_ICLRRuntimeInfo,(LPVOID *)&lpRuntimeInfo);
				if (FAILED(hrStart))
					return S_FALSE;
				hrStart = lpRuntimeInfo->GetInterface(CLSID_CLRRuntimeHost,IID_ICLRRuntimeHost,(LPVOID *)&m_pClrHost);
				if (FAILED(hrStart)) 
					return S_FALSE;

				hrStart = m_pClrHost->Start();
				if (FAILED(hrStart))
				{
					return S_FALSE;
				}
				if (hrStart == S_FALSE)
				{
					m_bCLRStart = true;
				}

				DWORD dwRetCode = 0;
				hrStart = m_pClrHost->ExecuteInDefaultAppDomain(m_strTangramCLRPath, _T("TangramCLR.Tangram"), _T("InitTangramCLR"), _T(""), &dwRetCode);
				dwRetCode = m_pMetaHost->Release();
				m_pMetaHost = nullptr;
				FreeLibrary(hMscoreeLib);
				if (hrStart != S_OK)
					return -1;
			}
		}
	}
	return 0;
}

typedef BOOL(WINAPI *IW64PFP)(HANDLE, BOOL *);
BOOL CTangramApp::IsWow64bit()
{
	BOOL bRet = false;
	IW64PFP  IW64P = (IW64PFP)GetProcAddress(GetModuleHandle(_T("kernel32")), "IsWow64Process");
	if (IW64P != NULL)
		IW64P(GetCurrentProcess(), &bRet);
	return bRet;
}

BOOL CTangramApp::Is64bitSystem()
{
	SYSTEM_INFO si;
	GetNativeSystemInfo(&si);

	if (si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64 ||si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_IA64)
		return true;
	else
		return false;
}

STDAPI DllCanUnloadNow(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return (AfxDllCanUnloadNow() == S_OK && theApp.GetLockCount() == 0) ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return theApp.DllGetClassObject(rclsid, riid, ppv);
}

STDAPI DllRegisterServer(void)
{
	if (IsWindows8Point1OrGreater())
	{
	}
	return theApp.DllRegisterServer();
}

STDAPI DllUnregisterServer(void)
{
	if (IsWindows8Point1OrGreater())
	{
	}
	return theApp.DllUnregisterServer();
}
