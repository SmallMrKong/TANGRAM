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

#include "stdafx.h"

#include "Wininet.h"
#include "TangramVSIApp.h"
#include "CommonIncludes.h"
//#include <Winbase.h>

CTangramApp theApp;

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
	char*   res;
	char   dwCode[20];
	DWORD   dwIndex, dwCodeLen;
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
			if (hFile)
				InternetCloseHandle(hFile);
			InternetCloseHandle(hSession);
			return   true;
		}
	}
	return   false;
}

BOOL CTangramApp::InitInstance()
{
	if (m_bInit == true)
		return true;
	m_bInit = true;
	TCHAR file[MAX_PATH];
	GetModuleFileName(::GetModuleHandle(_T("TangramPackage.dll")), file, MAX_PATH * sizeof(TCHAR));
	
	m_strModulePath = file;
	int nPos = m_strModulePath.ReverseFind('\\');
	m_strModulePath = m_strModulePath.Left(nPos + 1);
	HRESULT hr = SHGetFolderPath(NULL, CSIDL_APPDATA, NULL, 0, file);
	m_strAppDataPath = CString(file);
	m_strAppDataPath += _T("\\TangramData\\");

	hr = SHGetFolderPath(NULL, CSIDL_PROGRAM_FILES, NULL, 0, file);
	m_strProgramFilePath = CString(file);

	GetModuleFileName(NULL, file, MAX_PATH * sizeof(TCHAR));
	CString path(file);
	nPos = path.ReverseFind('\\');
	CString strName = path.Mid(nPos + 1);
	nPos = strName.Find(_T("."));
	m_strExeName = strName.Left(nPos);
	m_strAppDataPath += m_strExeName;
	m_strAppDataPath += _T("\\");
	m_strVersion = _T("");
	m_strDesignerXml = _T("");

	if (m_pDTE)
	{
		CComBSTR bstrVer(L"");
		m_pDTE->get_Version(&bstrVer);
		m_strVersion = OLE2T(bstrVer);
		m_strAppDataPath += m_strVersion;
		m_strAppDataPath += _T("\\");
	}
	
	CString strComponentURL = _T("");
	CString strActionXML = _T("");
	TCHAR szFile[MAX_PATH] = { 0 };
	::GetModuleFileName(NULL, szFile, MAX_PATH);
	CString strPath = CString(szFile);
	nPos = strPath.ReverseFind('\\');
	m_strPath = strPath.Left(nPos + 1);
	m_pTangram.CoCreateInstance(CComBSTR(L"tangram.tangram.1"));
	if (m_pTangram == NULL)
	{
		if (IsUserAdministrator() == false)
		{
			::MessageBox(NULL,_T("Please restart Visual Studio with Administrator AUTHORITY!"),_T("Visual Studio"),MB_OK);
			return false;
		}

		CString strXmlFile = m_strModulePath + _T("tangram.xml");
		if (::PathFileExists(strXmlFile))
		{
			CTangramXmlParse m_Parse;
			if (m_Parse.LoadFile(strXmlFile))
			{
				int nCount = 0;
				CTangramXmlParse* pParse = NULL;
				strComponentURL = m_Parse.attr(_T("ComponentURL"), _T(""));
				if (strComponentURL != _T(""))
				{
					pParse = m_Parse.GetChild(_T("InstallAction"));
					if (pParse)
						strActionXML = pParse->xml();
				}

				STARTUPINFO startupInfo;
				wchar_t cmdLine[2048];
				memset(&startupInfo, 0, sizeof(startupInfo));
				startupInfo.cb = sizeof(startupInfo);
				PROCESS_INFORMATION processInfo;
				memset(&processInfo, 0, sizeof(processInfo));
				DWORD code = 0;
				CTangramXmlParse* pXmlNode = m_Parse.GetChild(_T("MSI"));
				if (pXmlNode)
				{
					hr = SHGetFolderPath(NULL, CSIDL_WINDOWS, NULL, 0, szFile);
					m_strWindowSystemPath = CString(szFile);
					CString strSys64 = _T("\\SysWOW64\\");
					CString strPath = m_strWindowSystemPath + strSys64;
					CString strFile = _T("");
					if (::PathIsDirectory(strPath))
					{
						nCount = pXmlNode->GetCount();
						for (int i = 0; i < nCount; i++)
						{
							pParse = pXmlNode->GetChild(i);
							if (pParse)
							{
								strFile = theApp.m_strModulePath + _T("msi\\") + pParse->text();
								if (::PathFileExists(strFile))
								{
									swprintf_s(cmdLine, _T("msiexec.exe /i \"%s\" %s"), strFile.GetBuffer(), _T("/qb"));
									if (CreateProcess(NULL, cmdLine, NULL, NULL, true, 0, NULL, NULL, &startupInfo, &processInfo))
									{
										// wait for the installer to finish
										WaitForSingleObject(processInfo.hProcess, INFINITE);
										GetExitCodeProcess(processInfo.hProcess, &code);
									}
								}
							}
						}
					}
					else
					{
						strFile = theApp.m_strModulePath + _T("msi\\TangramCRTX86.msi");
						if (::PathFileExists(strFile))
						{
							swprintf_s(cmdLine, _T("msiexec.exe /i \"%s\" %s"), strFile.GetBuffer(), _T("/qb"));
							if (CreateProcess(NULL, cmdLine, NULL, NULL, true, 0, NULL, NULL, &startupInfo, &processInfo))
							{
								// wait for the installer to finish
								WaitForSingleObject(processInfo.hProcess, INFINITE);
								GetExitCodeProcess(processInfo.hProcess, &code);
							}
						}
					}
				}

				m_pTangram.CoCreateInstance(CComBSTR(L"tangram.tangram.1"));
				
				pXmlNode = m_Parse.GetChild(_T("COMLibs"));
				if (pXmlNode)
				{
					nCount = pXmlNode->GetCount();
					typedef int (CALLBACK* TANGRAMFUNCTION)(void);
					TANGRAMFUNCTION RegisterServerFunction = NULL;
					HINSTANCE handle = NULL;
					BOOL bRegisterServer = false;
					for (int i = 0; i < nCount; i++)
					{
						pParse = pXmlNode->GetChild(i);
						if (pParse)
						{
							CString strDll = _T("");
							CString strFile = pParse->text();
							CString strTarget = pParse->attr(_T("Target"), _T(""));
							nPos = strTarget.Find(_T("\\"));
							CString s = strTarget.Left(nPos);
							if (s.CompareNoCase(_T("Program Files")) == 0)
							{
								CString strDir = m_strProgramFilePath + strTarget.Mid(nPos);
								if (::PathIsDirectory(strDir) == false)
								{
									::SHCreateDirectoryEx(NULL, strDir, NULL);
								}
								CString strFile1 = strFile.Mid(strFile.Find(_T("\\")) + 1);
								strDll = strDir + strFile1;
								if (::PathFileExists(m_strModulePath + strFile))
								{
									::CopyFile(m_strModulePath + strFile, strDll, false);
								}
							}
							else
								strDll = m_strModulePath + strFile;
							if (::PathFileExists(strDll))
							{
								handle = LoadLibraryEx(strDll, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
								if (handle != NULL)
								{
									RegisterServerFunction = GetProcAddress(handle, "DllRegisterServer");
									if (RegisterServerFunction != NULL)
										bRegisterServer = (RegisterServerFunction() == S_OK);
									FreeLibrary(handle);
								}
							}
						}
					}
				}

				pXmlNode = m_Parse.GetChild(_T("CLRLibs"));
				if (pXmlNode)
				{
					CTangramXmlParse* pAssemblyBindingNode = NULL;
					CString strCfgFile = _T("");
					CString strVal = _T("");
					BOOL bModifyed = false;
					CTangramXmlParse m_Parse2;
					nCount = pXmlNode->GetCount();
					for (int i = 0; i < nCount; i++)
					{
						pParse = pXmlNode->GetChild(i);
						if (pParse)
						{
							CString _strLib = pParse->text();
							CString strLib = m_strModulePath + _strLib;
							if (::PathFileExists(strLib))
							{
								CString strTarget = theApp.m_strPath + _strLib;
								nPos = strTarget.ReverseFind('\\');
								if (nPos != -1)
								{
									CString strDir = strTarget.Left(nPos + 1);
									if (PathIsDirectory(strDir) == false)
									{
										::SHCreateDirectory(NULL, strDir);
										nPos = _strLib.ReverseFind('\\');
										CString strProbingPath = _strLib.Left(nPos);
										if (pAssemblyBindingNode == NULL&&strProbingPath!=_T(""))
										{
											TCHAR szFile[MAX_PATH] = { 0 };
											::GetModuleFileName(NULL, szFile, MAX_PATH);
											strCfgFile = CString(szFile) + _T(".config");
											if (m_Parse2.LoadFile(strCfgFile))
												pAssemblyBindingNode = m_Parse2[_T("runtime")][_T("assemblyBinding")].GetChild(_T("probing"));
											if (pAssemblyBindingNode)
											{
												if (strVal == _T(""))
													strVal = pAssemblyBindingNode->attr(_T("privatePath"), _T(""));
												CString _strval = strVal;
												_strval.MakeLower();
												CString strKey = _T(";") + strProbingPath + _T(";");
												strKey.MakeLower();
												if (_strval.Find(strKey) == -1)
												{
													bModifyed = true;
													nPos = strVal.Find(_T(";"));
													CString s = strVal.Left(nPos + 1);
													s += strProbingPath;
													s += _T(";");
													s += strVal.Mid(nPos);
													s.Replace(_T(";;"), _T(";"));
													strVal = s;
												}
											}
										}

									}
								}
								::CopyFile(strLib, strTarget, false);
							}
						}
					}
					if (bModifyed&&pAssemblyBindingNode)
					{
						pAssemblyBindingNode->put_attr(_T("privatePath"), strVal);
						m_Parse2.SaveFile(strCfgFile);
					}
				}

				pXmlNode = m_Parse.GetChild(_T("CommonFiles"));
				if (pXmlNode)
				{
					nCount = pXmlNode->GetCount();
					for (int i = 0; i < nCount; i++)
					{
						pParse = pXmlNode->GetChild(i);
						if (pParse)
						{
							CString strLib = m_strModulePath + pParse->text();
							if (::PathFileExists(strLib))
							{
								CString strTarget = m_strPath + strLib;
								nPos = strTarget.ReverseFind('\\');
								if (nPos != -1)
								{
									CString strDir = strTarget.Left(nPos + 1);
									if (PathIsDirectory(strDir) == false)
										::SHCreateDirectory(NULL, strDir);
								}
								::CopyFile(strLib, strTarget, false);
							}
						}
					}
				}

				pXmlNode = m_Parse.GetChild(_T("Exe"));
				if (pXmlNode)
				{
					nCount = pXmlNode->GetCount();
					for (int i = 0; i < nCount; i++)
					{
						pParse = pXmlNode->GetChild(i);
						if (pParse)
						{
							CString strLib = m_strModulePath + pParse->text();
							if (::PathFileExists(strLib))
							{
								swprintf_s(cmdLine, _T("%s /Regserver"), strLib.GetBuffer());
								if (CreateProcess(NULL, cmdLine, NULL, NULL, true, 0, NULL, NULL, &startupInfo, &processInfo))
								{
									// wait for the installer to finish
									WaitForSingleObject(processInfo.hProcess, INFINITE);
									GetExitCodeProcess(processInfo.hProcess, &code);
								}
							}
						}
					}
				}

				pXmlNode = m_Parse.GetChild(_T("OfficePlus"));
				if (pXmlNode)
				{
					//Install Office Component:
					BOOL m_b32bitOffice = false;
					BOOL m_b64bitSystem = false;

					SYSTEM_INFO si;
					GetNativeSystemInfo(&si);

					if (si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64 ||
						si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_IA64)
						m_b64bitSystem = true;

					wchar_t buf[_MAX_PATH];
					CString strPath = _T("SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration\\");
					swprintf_s(buf, strPath);
					DWORD size = _MAX_PATH;
					HKEY hKey = NULL;
					REGSAM hREGSAM = KEY_READ | KEY_QUERY_VALUE;
					if (m_b64bitSystem)
						hREGSAM |= KEY_WOW64_64KEY;
					if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, buf, 0, NULL, REG_OPTION_NON_VOLATILE, hREGSAM, NULL, &hKey, NULL) == ERROR_SUCCESS)
					{
						wchar_t szPlatForm[_MAX_PATH];
						if (RegQueryValueEx(hKey, L"Platform", 0, NULL, (BYTE*)&szPlatForm, &size) == ERROR_SUCCESS)
						{
							CString strVal = path;
							if (strVal.CompareNoCase(_T("x86")) == 0)
								m_b32bitOffice = true;
						}
						RegCloseKey(hKey);
					}
					CString strOfficePath = _T("");
					strPath = _T("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\");
					strPath += _T("winword.exe");
					swprintf_s(buf, strPath);
					hKey = NULL;
					if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, buf, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_READ | KEY_QUERY_VALUE | KEY_WOW64_64KEY, NULL, &hKey, NULL) == ERROR_SUCCESS)
					{
						wchar_t path[_MAX_PATH];
						if (RegQueryValueEx(hKey, L"Path", 0, NULL, (BYTE*)&path, &size) == ERROR_SUCCESS)
						{
							strOfficePath = path;
						}
						RegCloseKey(hKey);
					}

					CTangramXmlParse* pXmlNode2 = pXmlNode->GetChild(_T("ConfigFile"));
					if (pXmlNode2)
					{
						nCount = pXmlNode2->GetCount();
						for (int i = 0; i < nCount; i++)
						{
							pParse = pXmlNode2->GetChild(i);
							if (pParse)
							{
								CString strCfgFile = m_strModulePath + _T("OfficePlus\\") + pParse->text();
								CString strTarget = strOfficePath + pParse->text();
								if (::PathFileExists(strCfgFile))
								{
									::CopyFile(strCfgFile, strTarget, false);
								}
							}
						}
					}
					if (m_b32bitOffice)
					{
						//for 32bit Office:
						pXmlNode2 = pXmlNode->GetChild(_T("x86"));
					}
					else
					{
						//for 64bit Office:
						pXmlNode2 = pXmlNode->GetChild(_T("x64"));
					}
					if (pXmlNode2)
					{
						nCount = pXmlNode2->GetCount();
						for (int i = 0; i < nCount; i++)
						{
							pParse = pXmlNode2->GetChild(i);
							if (pParse)
							{
								CString strSourceFile = m_strModulePath + _T("OfficePlus\\") + pXmlNode2->name() + _T("\\") + pParse->text();
								CString strTarget = strOfficePath + pParse->name() +_T("\\") + pParse->text();
								if (::PathFileExists(strSourceFile))
								{
									::CopyFile(strSourceFile, strTarget, false);
								}
							}
						}
					}
				}
			}
		}
	}

	if (m_pTangram)
	{
		CTangramPackageProxy* pProxy = (CTangramPackageProxy*)m_pTangramPackage;
		m_pTangram->put_AppKeyValue(CComBSTR(L"TangramPackageProxy"), CComVariant((LONGLONG)pProxy));
		m_pTangram->put_AppExtender(CComBSTR(L"DTE"), m_pDTE);
		long hMainWnd = 0;
		CComPtr<VxDTE::Window> pMainWnd;
		m_pDTE->get_MainWindow(&pMainWnd);
		if (pMainWnd)
		{
			pMainWnd->get_HWnd(&hMainWnd);
			if (hMainWnd)
			{
				m_pTangram->put_AppKeyValue(CComBSTR(L"MainWnd"), CComVariant((LONG)hMainWnd));
			}
		}
		m_strDesignerXml = _T("<TangramDesigner><window><node name=\"TangramDesigner\" id =\"ActiveX\" cnnid=\"about:blank\" /></window></TangramDesigner>");
		//VARIANT v;
		//m_pTangram->get_AppKeyValue(CComBSTR(L"tangramdesigner"), &v);
		//if (v.vt = VT_BSTR)
		//{
		//	m_strDesignerXml = OLE2T(v.bstrVal);
		//}
		//::VariantClear(&v);
		CComQIPtr<DTE2> pDTE2(m_pDTE);
		CComPtr<ToolWindows> pToolWindows;
		pDTE2->get_ToolWindows(&pToolWindows);
		if (pToolWindows)
		{
			CComPtr<VxDTE::OutputWindow> pOutWnd;
			pToolWindows->get_OutputWindow(&pOutWnd);
			if (pOutWnd)
			{
				CComPtr<OutputWindowPane> pPane;
				CComPtr<OutputWindowPanes> pPanes;
				pOutWnd->get_OutputWindowPanes(&pPanes);
				pPanes->Item(CComVariant(L"Tangram"), &pPane);
				if (pPane == NULL)
				{
					pPanes->Add(CComBSTR(L"Tangram"), &pPane);
					pPane->OutputString(CComBSTR(L"Welcome to Tangram for Visual Studio!\r\n"));
				}
				m_pOutputWindowPane = pPane.Detach();
				m_pOutputWindowPane->AddRef();
			}
		}
	}
	return true;
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

// This macro is used as default registry root when a NULL parameter is passed to VSDllRegisterServer
// or VSDllUnregisterServer. For sample code we set as default the experimental instance, but for production
// code you should change it to the standard VisualStudio instance that is LREGKEY_VISUALSTUDIOROOT.
#define DEFAULT_REGISTRY_ROOT LREGKEY_VISUALSTUDIOROOT

// Since this project defines an oleautomation interface, the typelib needs to be registered.
#define VSL_REGISTER_TYPE_LIB true

// Must come after declaration of _AtlModule and DEFAULT_REGISTRY_ROOT
#ifndef VSL_REGISTER_TYPE_LIB
#define VSL_REGISTER_TYPE_LIB false
#endif VSL_REGISTER_TYPE_LIB

#pragma warning(push) // Sometimes true, sometimes not.
#pragma warning(disable : 4702) // warning C4702: unreachable code

// Initializes ATL
extern "C"
BOOL WINAPI DllMain(HINSTANCE /*hInstance*/, DWORD dwReason, LPVOID lpReserved)
{
	VSL_STDMETHODTRY{

		return theApp.DllMain(dwReason, lpReserved);

	}VSL_STDMETHODCATCH()

		return false;
}

// Used by COM to determine whether the DLL can be unloaded
STDAPI DllCanUnloadNow()
{
	VSL_STDMETHODTRY{

		return theApp.DllCanUnloadNow();

	}VSL_STDMETHODCATCH()

		return VSL_GET_STDMETHOD_HRESULT();
}

// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	VSL_STDMETHODTRY{

		return theApp.GetClassObject(rclsid, riid, ppv);

	}VSL_STDMETHODCATCH()

		return VSL_GET_STDMETHOD_HRESULT();
}


STDMETHODIMP VSDllRegisterServerInternal(_In_opt_ wchar_t* strRegRoot, bool shouldRegister, bool isRanu)
{
	VSL_STDMETHODTRY{

		VsRegistryUtilities::SetRegRoot(NULL == strRegRoot ? DEFAULT_REGISTRY_ROOT : strRegRoot, isRanu);

	//Set ATL to register the typelib as RANU if requested by the caller.
	AtlSetPerUserRegistration(isRanu);

	if (shouldRegister)
	{
		return theApp.RegisterServer(VSL_REGISTER_TYPE_LIB);
	}
	else
	{
		HRESULT hr = theApp.UnregisterServer(VSL_REGISTER_TYPE_LIB);

		// If the type library was already unregistered, ignore the failure
		return (TYPE_E_REGISTRYACCESS == hr) ? S_OK : hr;
	}
	}VSL_STDMETHODCATCH()

		return VSL_GET_STDMETHOD_HRESULT();
}

// Registers COM objects normally and registers VS Packages to the specified VS registry hive under HKCU
STDAPI VSDllRegisterServerUser(_In_opt_ wchar_t* strRegRoot)
{
	return VSDllRegisterServerInternal(strRegRoot, true, true);
}

// Unregisters COM objects normally and unregisters VS Packages from the specified VS registry hive under HKCU
STDAPI VSDllUnregisterServerUser(__in_opt wchar_t* strRegRoot)
{
	return VSDllRegisterServerInternal(strRegRoot, false, true);
}

// Registers COM objects normally and registers VS Packages to the specified VS registry hive
STDAPI VSDllRegisterServer(__in_opt wchar_t* strRegRoot)
{
	return VSDllRegisterServerInternal(strRegRoot, true, false);
}

// Unregisters COM objects normally and unregisters VS Packages from the specified VS registry hive
STDAPI VSDllUnregisterServer(__in_opt wchar_t* strRegRoot)
{
	return VSDllRegisterServerInternal(strRegRoot, false, false);
}

// Registers COM objects normally and registers VS Packages to the default VS registry hive
STDAPI DllRegisterServer()
{
	return VSDllRegisterServer(NULL);
}

// Unregisters COM objects normally and unregisters VS Packages from the default VS registry hive
STDAPI DllUnregisterServer()
{
	return VSDllUnregisterServer(NULL);
}

#pragma warning(pop)

