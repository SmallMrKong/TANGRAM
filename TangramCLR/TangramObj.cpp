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

#include "StdAfx.h"
#include "dllmain.h"
#include "TangramNodeCLREvent.h"
#include "TangramObj.h"

using namespace System::Threading;
using namespace System::Diagnostics;
using namespace System::Reflection;

namespace TangramCLR
{
	WndNode^ WndNodeCollection::default::get(int iIndex)
	{
		WndNode^ Node = nullptr;
		IWndNode* pNode = nullptr;
		m_pNodeCollection->get_Item(iIndex,&pNode);
		return theAppProxy._createObject<IWndNode, WndNode>(pNode);
	}

	WndNode::WndNode(IWndNode* pNode)
	{
		LONGLONG m_nConnector = 0;
		m_hWnd = NULL;
		m_pTangramNodeEvent = new CTangramNodeEvent();
		m_pTangramNodeEvent->m_pWndNode = pNode;
		m_pTangramNodeEvent->m_pTangramNodeCLREvent = new CWndNodeCLREvent();
		m_pTangramNodeEvent->m_pTangramNodeCLREvent->m_pWndNode = this;
		HRESULT hr = m_pTangramNodeEvent->DispEventAdvise(pNode);

		m_pChildNodes = nullptr;
		SetNewNode(pNode);
		LONGLONG nValue = (LONGLONG)pNode;
		theAppProxy._insertObject(nValue, this);
	}

	WndNode::~WndNode()
	{
		SetNewNode(NULL);
	}

	WndNodeCollection^ WndNode::ChildNodes::get()
	{
		if (m_pChildNodes == nullptr)
		{
			CComPtr<IWndNodeCollection> pNodes;

			m_pWndNode->get_ChildNodes(&pNodes);
			m_pChildNodes = gcnew WndNodeCollection(pNodes);
		}
		return m_pChildNodes;
	}

	WndPage^ WndNode::WndPage::get()
	{
		IWndPage* pPage = nullptr;
		m_pWndNode->get_WndPage(&pPage);

		if (pPage)
		{
			return theAppProxy._createObject<IWndPage, TangramCLR::WndPage>(pPage);
		}
		return nullptr;
	}

	String^ WndNode::Caption::get()
	{
		if(m_pWndNode)
		{
			CComBSTR bstrCap("");
			m_pWndNode->get_Caption(&bstrCap);
			String^ strCap = Marshal::PtrToStringUni((System::IntPtr)LPTSTR(LPCTSTR(bstrCap)));
			return strCap;
		}
		return "";
	}

	void WndNode::Init()
	{
		long h = 0;
		if (m_pWndNode)
		{ 
			m_pWndNode->get_Handle(&h);
			::SendMessage((HWND)h, WM_TANGRAMMSG, 1, 0);
		}
	}

	Object^ WndNode::PlugIn::get(String^ strObjName)
	{
		Object^ pObj = nullptr;
		if(m_pWndNode)
		{
			WndNode^ pRootNode = this->RootNode;
			if (pRootNode->m_pTangramPlugInDic == nullptr)
			{
				pRootNode->m_pTangramPlugInDic = gcnew Dictionary<String^, Object^>();
			}
			if (pRootNode->m_pTangramPlugInDic->TryGetValue(strObjName, pObj) == false)
			{
				IDispatch* pDisp = nullptr;
				LRESULT hr = m_pWndNode->get_AxPlugIn(STRING2BSTR(strObjName),&pDisp);
				if(SUCCEEDED(hr)&&pDisp)
				{
					Object^ pObj = reinterpret_cast<Object^>(Marshal::GetObjectForIUnknown((System::IntPtr)(pDisp)));
					pRootNode->m_pTangramPlugInDic[strObjName] = pObj;
					return pObj;
				}

			}
		}
		return pObj;
	}

	Tangram::Tangram()
	{
		m_nMDIClientHandle = (IntPtr)0;
		theApp.m_pCloudAddinCLREvent = new CTangramCoreCLREvent();
		m_pTangramInitGlobalTask = gcnew Func<Object^, String^>(&TangramCLR::Tangram::TangramGlobalFunc);
	}

	Tangram::~Tangram(void)
	{
		for each (KeyValuePair<String^, Object^>^ dic in TangramCLR::Tangram::m_pTangramCLRObjDic)
		{
			if (dic->Value != nullptr)
			{
				Object^ cmd = dic->Value;
				if (dic->Key != L"HttpApplication")
					delete cmd;
			}
		}
	}

	String^ Tangram::ComputeHash(String^ source)
	{
		BSTR bstrSRC = STRING2BSTR(source);
		LPCWSTR srcInfo = OLE2T(bstrSRC);
		std::string strSrc = (LPCSTR)CW2A(srcInfo, CP_UTF8);
		int nSrcLen = strSrc.length();
		CString strRet = _T("");
		theApp.CalculateByteMD5((BYTE*)strSrc.c_str(), nSrcLen, strRet);
		::SysFreeString(bstrSRC);
		return BSTR2STRING(strRet);
	}

	Object^ Tangram::Application::get()
	{
		Object^ pRetObject = nullptr;
		if (theApp.m_pTangram)
		{
			try
			{
				IDispatch* pApp = nullptr;
				theApp.m_pTangram->get_Application(&pApp);
				if (pApp)
				{
					pRetObject = Marshal::GetObjectForIUnknown((System::IntPtr)pApp);
				}
			}
			catch (InvalidOleVariantTypeException^e)
			{

			}
			catch (NotSupportedException^ e)
			{

			}
		}
		return pRetObject;
	}

	void Tangram::Application::set(Object^ obj)
	{
		if (theApp.m_pTangram)
		{
			try
			{
				VARIANT var;
				Marshal::GetNativeVariantForObject(obj, (System::IntPtr)&var);
				theApp.m_pTangram->put_ExternalInfo(var);
			}
			catch (ArgumentException^ e)
			{

			}
		}
	}

	WndNode^ Tangram::CreatingNode::get()
	{
		Object^ pRetObject = nullptr;
		if (theApp.m_pTangram)
		{
			IWndNode* pNode = nullptr;
			theApp.m_pTangram->get_CreatingNode(&pNode);
			if (pNode)
				return theAppProxy._createObject<IWndNode, WndNode>(pNode);
		}
		return nullptr;
	}

	String^ Tangram::AppKeyValue::get(String^ iIndex)
	{
		CComVariant bstrVal(::SysAllocString(L""));
		theApp.m_pTangram->get_AppKeyValue(STRING2BSTR(iIndex), &bstrVal);
		String^ strVal = BSTR2STRING(bstrVal.bstrVal);
		//::SysFreeString(bstrVal);
		return strVal;
	}

	void Tangram::AppKeyValue::set(String^ iIndex, String^ newVal)
	{
		theApp.m_pTangram->put_AppKeyValue(STRING2BSTR(iIndex), CComVariant(STRING2BSTR(newVal)));
	}

	void Tangram::Fire_OnClose()
	{
		OnClose();
	}

	Dictionary<String^, WndNode^>^ Tangram::InitObj(Control^ pObj , String^ strXML)
	{
		if (pObj == nullptr)
			return nullptr;
		String^ strRes = L"";
		if (String::IsNullOrEmpty(strXML))
		{
			strRes = GetTangramRes(pObj, "");
		}
		else
			strRes = strXML;
		CString strXml = strRes;
		CTangramXmlParse m_Parse;
		BOOL bXml = m_Parse.LoadXml(strXml);
		if (bXml == false)
		{
			bXml = m_Parse.LoadFile(strXml);
			if (!bXml)
			{
				return nullptr;
			}
		}
		CString strName = pObj->Name;
		CString strTypeName = pObj->GetType()->Name;
		CTangramXmlParse* pFormsParse = m_Parse.GetChild(_T("Forms"));
		if(pFormsParse==nullptr)
			pFormsParse = m_Parse.GetChild(_T("forms"));
		if (pFormsParse)
		{
			CTangramXmlParse* pFormParse = pFormsParse->GetChild(strName);
			if (pFormParse == nullptr)
			{
				pFormParse = pFormsParse->GetChild(strName.MakeLower());
				if (pFormParse == nullptr)
				{
					pFormParse = pFormsParse->GetChild(strTypeName);
					if (pFormParse == nullptr)
					{
						pFormParse = pFormsParse->GetChild(strTypeName.MakeLower());
					}
				}
			}
			if (pFormParse)
			{
				Dictionary<String^, WndNode^>^ pDictionary = nullptr;
				int nCount = pFormParse->GetCount();
				for (int i = 0; i < nCount; i++)
				{
					CTangramXmlParse* pChild = pFormParse->GetChild(i);
					if (pChild)
					{
						Control^ _ctrl = nullptr;
						CString strCtrlName = pChild->name();
						String^ _ctrlName = BSTR2STRING(strCtrlName);
						if (strCtrlName == _T("mdiclient"))
						{
							_ctrl = GetMDIClient((Form^)pObj);
						}
						else
						{
							cli::array<Control^, 1>^ pArray = pObj->Controls->Find(_ctrlName, true);
							if (pArray != nullptr&&pArray->Length)
								_ctrl = pArray[0];
						}
						if (_ctrl != nullptr)
						{
							long hHost = _ctrl->Handle.ToInt64();
							long hParent = _ctrl->Parent->Handle.ToInt64();
							IWndPage* pPage = nullptr;
							theApp.m_pTangram->CreateWndPage(hParent, &pPage);
							if (pPage)
							{
								IWndFrame* pFrame = nullptr;
								pPage->CreateFrame(CComVariant(0), CComVariant((long)hHost), CComBSTR(strCtrlName), &pFrame);
								if (pFrame)
								{
									IWndNode* pNode = nullptr;
									pFrame->Extend(CComBSTR(L"Default"), pChild->xml().AllocSysString(), &pNode);
									if (pNode)
									{
										if (pDictionary == nullptr)
											pDictionary = gcnew Dictionary<String^,WndNode^>;
										pDictionary->Add(_ctrlName, theAppProxy._createObject<IWndNode, WndNode>(pNode));
									}
								}
							}
						}
					}
				}
				return pDictionary;
			}
		}
		return nullptr;
	}

	int Tangram::InitTangramCLR(String^ pwzArgument)
	{
		return 0;
	}

	Control^ Tangram::GetMDIClient(Form^ pForm)
	{	
		if (pForm&&pForm->IsMdiContainer)
		{
			int nCount = pForm->Controls->Count;
			String^ strName = L"";
			for (int i = nCount - 1; i >= 0; i--)
			{
				Control^ pCtrl = pForm->Controls[i];
				strName = pCtrl->GetType()->Name->ToLower();
				if (strName == L"mdiclient")
				{
					return pCtrl;
				}
			}
		}
		return nullptr;
	}

	String^ Tangram::TangramGlobalFunc(Object^ _pThisObj)
	{
		return L"";
	};

	Object^ Tangram::ActiveMethod(String^ strObjID, String^ strMethod, cli::array<Object^, 1>^ p)
	{
		Object^ pRetObj = nullptr;
		Tangram^ pApp = Tangram::GetTangram();
		String^ strIndex = strObjID + L"|" + strMethod;
		MethodInfo^ mi = nullptr;
		Object^ pObj = nullptr;
		if (pApp->m_pTangramCLRMethodDic->TryGetValue(strIndex, mi) == true)
		{
			try
			{
				pRetObj = mi->Invoke(pObj, p);
			}
			finally
			{
			}
			return pRetObj;
		}

		if (pApp->m_pTangramCLRObjDic->TryGetValue(strObjID, pObj) == false)
		{
			pObj = CreateObject(strObjID);
			pApp->m_pTangramCLRObjDic[strObjID] = pObj;
		}
		if (pObj != nullptr)
		{
			MethodInfo^ mi = nullptr;
			try
			{
				mi = pObj->GetType()->GetMethod(strMethod);
				pApp->m_pTangramCLRMethodDic[strIndex] = mi;
			}
			catch (AmbiguousMatchException^ e)
			{
				Debug::WriteLine(L"Tangram::ActiveMethod GetMethod: " + e->Message);
			}
			catch (ArgumentNullException^ e)
			{
				Debug::WriteLine(L"Tangram::ActiveMethod GetMethod: " + e->Message);
			}
			finally
			{
				if (mi != nullptr)
				{
					try
					{
						pRetObj = mi->Invoke(pObj, p);
					}
					finally
					{
					}
				}
			}
		}

		return pRetObj;
	}
			
	Object^ Tangram::ActiveObjectMethod(Object^ pObj, String^ strMethod, cli::array<Object^, 1>^ p)
	{
		Object^ pRetObj = nullptr;

		if (pObj != nullptr)
		{
			MethodInfo^ mi = nullptr;
			try
			{
				mi = pObj->GetType()->GetMethod(strMethod);
			}
			catch (AmbiguousMatchException^ e)
			{
				Debug::WriteLine(L"Tangram::ActiveMethod GetMethod: " + e->Message);
			}
			catch (ArgumentNullException^ e)
			{
				Debug::WriteLine(L"Tangram::ActiveMethod GetMethod: " + e->Message);
			}
			finally
			{
				if (mi != nullptr)
				{
					try
					{
						pRetObj = mi->Invoke(pObj, p);
					}
					finally
					{
					}
				}
			}
		}

		return pRetObj;
	}

	WndPage^ Tangram::CreateWndPage(Control^ ctrl, Object^ ExternalObj)
	{
		if (ctrl != nullptr)
		{
			LONGLONG hWnd = ctrl->Handle.ToInt64();
			IWndPage* pPage = nullptr;
			theApp.m_pTangram->CreateWndPage(hWnd, &pPage);
			if (pPage)
			{
				WndPage^ _pTangram =  theAppProxy._createObject<IWndPage, WndPage>(pPage);
				if (_pTangram != nullptr&&ExternalObj != nullptr)
				{
					_pTangram->External = ExternalObj; 
					return _pTangram;
				}
			}
		}
		return nullptr;
	}

	Object^ Tangram::CreateObject(String^ strObjID)
	{
		if (String::IsNullOrEmpty(strObjID) == true)
			return nullptr;
		String^ m_strID = strObjID->ToLower();
		String^ strLib = nullptr;
		Object^ m_pObj = nullptr;
		if (m_strID != L"")
		{
			Type^ pType = nullptr;
			Tangram^ pApp = Tangram::GetTangram();
			Monitor::Enter(pApp->m_pTangramCLRTypeDic);
			String^ strID = nullptr;
			if (pApp->m_pTangramCLRTypeDic->TryGetValue(m_strID, pType) == false)
			{
				Assembly^ m_pDotNetAssembly = nullptr;
				bool bSystemObj = false;
				int nIndex = m_strID->IndexOf(L",");
				strID = m_strID->Substring(0, nIndex);
				strLib = m_strID->Substring(nIndex + 1, m_strID->Length - nIndex - 1);
				nIndex = m_strID->IndexOf(L"system.windows.forms");
				if (nIndex != -1)
				{
					bSystemObj = true;
					Control^ pObj = gcnew Control();
					m_pDotNetAssembly = pObj->GetType()->Assembly;
				}
				else
				{
					try
					{
						m_pDotNetAssembly = Assembly::Load(strLib);
					}
					catch (ArgumentNullException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject: " + e->Message);
					}
					catch (ArgumentException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject: " + e->Message);
					}
					catch (FileNotFoundException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject: " + e->Message);
					}
					catch (FileLoadException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject: " + e->Message);
					}
					catch (BadImageFormatException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject: " + e->Message);
					}
				}
				if (m_pDotNetAssembly != nullptr)
				{
					try
					{
						pType = m_pDotNetAssembly->GetType(strID, true, true);
					}
					catch (TypeLoadException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject Assembly->GetType: " + e->Message);
					}
					catch (ArgumentNullException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject Assembly->GetType: " + e->Message);
					}
					catch (ArgumentException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject Assembly->GetType: " + e->Message);
					}
					catch (FileNotFoundException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject Assembly->GetType: " + e->Message);
					}
					catch (FileLoadException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject Assembly->GetType: " + e->Message);
					}
					catch (BadImageFormatException^ e)
					{
						Debug::WriteLine(L"Tangram CreateObject Assembly->GetType: " + e->Message);
					}
					finally
					{
						if (pType != nullptr)
						{
							pApp->m_pTangramCLRTypeDic->Add(m_strID, pType);
						}
						if (m_pDotNetAssembly != nullptr)
							delete m_pDotNetAssembly;
					}
				}
			}
			Monitor::Exit(pApp->m_pTangramCLRTypeDic);
			if (pType)
			{
				try
				{
					m_pObj = (Object^)Activator::CreateInstance(pType);
				}
				catch (TypeLoadException^ e)
				{
					Debug::WriteLine(L"Tangram CreateObject Activator::CreateInstance: " + e->Message);
				}
				catch (ArgumentNullException^ e)
				{
					Debug::WriteLine(L"Tangram CreateObject Activator::CreateInstance: " + e->Message);
				}
				catch (ArgumentException^ e)
				{
					Debug::WriteLine(L"Tangram CreateObject Activator::CreateInstance: " + e->Message);
				}
				catch (NotSupportedException^ e)
				{
					Debug::WriteLine(L"Tangram CreateObject Activator::CreateInstance: " + e->Message);
				}
				catch (TargetInvocationException^ e)
				{
					Debug::WriteLine(L"Tangram CreateObject Activator::CreateInstance: " + e->Message);
				}
				catch (MethodAccessException^ e)
				{
					Debug::WriteLine(L"Tangram CreateObject Activator::CreateInstance: " + e->Message);
				}
				catch (InvalidComObjectException^ e)
				{
					Debug::WriteLine(L"Tangram CreateObject Activator::CreateInstance: " + e->Message);
				}
				catch (MissingMethodException^ e)
				{
					Debug::WriteLine(L"Tangram CreateObject Activator::CreateInstance: " + e->Message);
				}
				catch (COMException^ e)
				{
					Debug::WriteLine(L"Tangram CreateObject Activator::CreateInstance: " + e->Message);
				}
			}
		}
		return m_pObj;
	}

	WndNode^ WndNode::Extend(String^ strKey, String^ strXml)
	{
		if (m_pWndNode)
		{
			BSTR bstrKey = STRING2BSTR(strKey);
			BSTR bstrXml = STRING2BSTR(strXml);
			IWndNode* pNode = nullptr;
			m_pWndNode->Extend(bstrKey, bstrXml, &pNode);
			::SysFreeString(bstrKey);
			::SysFreeString(bstrXml);
			if (pNode)
			{
				return theAppProxy._createObject<IWndNode, WndNode>(pNode);
			}
		}
		return nullptr;
	}

	Object^ WndNode::ActiveMethod(String^ strMethod, cli::array<Object^, 1>^ p)
	{
		Object^ pRetObj = nullptr;
		if (m_pHostObj != nullptr)
		{
			MethodInfo^ mi = nullptr;
			if (m_pTangramCLRMethodDic==nullptr)
				m_pTangramCLRMethodDic = gcnew Dictionary<String^, MethodInfo^>();
			Object^ pObj = nullptr;
			if (m_pTangramCLRMethodDic->TryGetValue(strMethod, mi) == true)
			{
				try
				{
					pRetObj = mi->Invoke(m_pHostObj, p);
				}
				finally
				{
				}
				return pRetObj;
			}
			try
			{
				mi = m_pHostObj->GetType()->GetMethod(strMethod);
				m_pTangramCLRMethodDic[strMethod] = mi;
			}
			catch (AmbiguousMatchException^ e)
			{
				Debug::WriteLine(L"Tangram::ActiveMethod GetMethod: " + e->Message);
			}
			catch (ArgumentNullException^ e)
			{
				Debug::WriteLine(L"Tangram::ActiveMethod GetMethod: " + e->Message);
			}
			finally
			{
				if (mi != nullptr)
				{
					try
					{
						pRetObj = mi->Invoke(m_pHostObj, p);
					}
					finally
					{
					}
				}
			}
		}

		return pRetObj;
	}

	//TangramCLR::WndNode^ TangramCLR::CloudAddin::ExtendCtrl(Control^ pCtrl, String^ strFile)
	//{
	//	if(pCtrl==nullptr||String::IsNullOrEmpty(strFile))
	//		return nullptr;
	//	if(pCtrl->Parent==nullptr)
	//		return nullptr;

	//	CComBSTR m_strFile;
	//	IntPtr ip = Marshal::StringToBSTR(strFile);
	//	m_strFile = static_cast<BSTR>(ip.ToPointer());
	//	Marshal::FreeBSTR(ip);

	//	CComPtr<IWndNode> pDocObj;
	//	theApp.m_pTangram->Extend((LONGLONG)pCtrl->Handle.ToInt64(), m_strFile, &pDocObj);
	//	WndNode^ retNode = theAppProxy._createObject<IWndNode, WndNode>(pDocObj);
	//	return retNode;
	//}

	//Object^ TangramCLR::TangramNode::ExecScript(String^ strScript)
	//{
	//	if (String::IsNullOrEmpty(strScript) == false)
	//	{
	//		CComVariant v;
	//		m_pWndNode->ExecScript(STRING2BSTR(strScript), &v);
	//		Object^ ret =  Marshal::GetObjectForNativeVariant((IntPtr)&v);
	//		return ret;
	//	}
	//	return nullptr;
	//}

	WndPage::WndPage(void)
	{
	}

	WndPage::WndPage(IWndPage* pPage)
	{
		m_pPage = pPage;
		LONGLONG nValue = (LONGLONG)m_pPage;
		theAppProxy._insertObject(nValue, this);
		m_pTangramClrEvent = new CWndPageClrEvent();
		m_pTangramClrEvent->DispEventAdvise(m_pPage);
		m_pTangramClrEvent->m_pPage = this;
	}

	WndPage::~WndPage()
	{
		m_pTangramClrEvent->DispEventUnadvise(m_pPage);
		LONGLONG nValue = (LONGLONG)m_pPage;
		theAppProxy._removeObject(nValue);
		delete m_pTangramClrEvent;
	}

	WndNode^ WndPage::GetWndNode(String^ strXml, String^ strFrameName)
	{
		if (String::IsNullOrEmpty(strXml) || String::IsNullOrEmpty(strFrameName))
			return nullptr;
		BSTR bstrXml = STRING2BSTR(strXml);
		BSTR bstrFrameName = STRING2BSTR(strFrameName);
		CComPtr<IWndNode> pNode;
		m_pPage->GetWndNode(bstrXml, bstrFrameName, &pNode);
		WndNode^ pRetNode = nullptr;
		if (pNode)
		{
			pRetNode = theAppProxy._createObject<IWndNode, WndNode>(pNode);
		}
		::SysFreeString(bstrXml);
		::SysFreeString(bstrFrameName);
		return pRetNode;
	}

	//void Tangram::ExecScript(String^ strScript)
	//{
	//	if (String::IsNullOrEmpty(strScript))
	//		return;
	//	BSTR bstrScript = STRING2BSTR(strScript);
	//	m_pPage->ExecScript(bstrScript);
	//	::SysFreeString(bstrScript);
	//}

	void WndPage::AddObjToHtml(String^ strObjName, bool bConnectEvent, Object^ pObj)
	{
		if (String::IsNullOrEmpty(strObjName) || pObj == nullptr)
			return;

		BSTR bstrName = STRING2BSTR(strObjName);
		m_pPage->AddObjToHtml(bstrName, bConnectEvent, (IDispatch*)Marshal::GetIUnknownForObject(pObj).ToPointer());
		::SysFreeString(bstrName);
	}

	WndNode^ WndFrame::ExtendXml(String^ strKeyName,String^ strXml)
	{
		if (String::IsNullOrEmpty(strXml) || String::IsNullOrEmpty(strKeyName))
			return nullptr;
		BSTR bstrXml = STRING2BSTR(strXml);
		BSTR bstrKey = STRING2BSTR(strKeyName);
		CComPtr<IWndNode> pNode;
		m_pWndFrame->Extend(bstrKey,bstrXml, &pNode);
		WndNode^ pRetNode = nullptr;
		if (pNode)
		{
			pRetNode = theAppProxy._createObject<IWndNode, WndNode>(pNode);
		}
		::SysFreeString(bstrXml);
		::SysFreeString(bstrKey);
		return pRetNode;
	}

	Object^ WndFrame::FrameData::get(String^ iIndex)
	{
		CComVariant bstrVal(::SysAllocString(L""));
		m_pWndFrame->get_FrameData(STRING2BSTR(iIndex), &bstrVal);
		return Marshal::GetObjectForNativeVariant((IntPtr)&bstrVal);;
	}

	void WndFrame::FrameData::set(String^ iIndex, Object^ newVal)
	{
		IntPtr nPtr = (IntPtr)0;
		Marshal::GetNativeVariantForObject(newVal, nPtr);
		m_pWndFrame->put_FrameData(STRING2BSTR(iIndex), *(VARIANT*)nPtr.ToInt64());
	}
}