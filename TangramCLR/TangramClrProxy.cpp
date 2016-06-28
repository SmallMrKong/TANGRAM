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

// HostExtender.cpp : Implementation of CTangramNavigator

#include "stdafx.h"
#include "dllmain.h" 
#include "TangramClrProxy.h"
#include "TangramNodeCLREvent.h"

using namespace TangramCLR;

typedef HRESULT (__stdcall *TangramCLRCreateInstance)(REFCLSID clsid, REFIID riid, /*iid_is(riid)*/ LPVOID *ppInterface);

CCloudAddinCLRProxy theAppProxy;

CCloudAddinCLRProxy::CCloudAddinCLRProxy()
{
	Application::EnableVisualStyles();
	m_pPropertyGrid = nullptr;
	m_pVSPropertyGrid = nullptr;
	m_pSystemAssembly = nullptr;
	m_htObjects = gcnew Hashtable();
	CComPtr<ITangram> pCloudAddin;
	pCloudAddin.CoCreateInstance(CComBSTR("Tangram.Tangram.1"));
	theApp.m_pTangram = pCloudAddin.Detach();
	if (theApp.m_pTangram)
	{
		if(theApp.DispEventAdvise(theApp.m_pTangram)==S_OK)
			theApp.m_pTangram->put_AppKeyValue(CComBSTR(L"CLRProxy"), CComVariant((LONGLONG)this));
	}
	m_pTangramProxy = gcnew TangramProxy();
	m_hCurDesignObj = NULL;
}

CCloudAddinCLRProxy::~CCloudAddinCLRProxy()
{
	if(TangramCLR::Tangram::m_pAppQuitEvent !=nullptr)
		TangramCLR::Tangram::m_pAppQuitEvent->WaitOne();
	if (m_strCollaborationScript!=_T(""))
	{
	}
	TangramCLR::Tangram^ pManager = TangramCLR::Tangram::GetTangram();
	delete pManager;
	if (m_pTangramProxy)
	{
		delete m_pTangramProxy;
	}
}

BSTR CCloudAddinCLRProxy::AttachObjEvent(IDispatch* EventObj, IDispatch* SourceObj, WindowEventType EventType, IDispatch* HTMLWindow)
{
	BSTR bstrRet = L"";
	Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)SourceObj);
	if (pCtrl != nullptr)
	{
		bstrRet = STRING2BSTR(pCtrl->Name);
		LONGLONG nIndex = (LONGLONG)HTMLWindow;
		HWND hWnd = (HWND)pCtrl->Handle.ToInt64();
		ObjectEventInfo* pInfo = (ObjectEventInfo*)::GetWindowLongPtr(hWnd, GWLP_USERDATA);
		if (pInfo == nullptr)
		{
			m_pTangramProxy->AttachHandleDestroyedEvent(pCtrl);
			pInfo = new ObjectEventInfo();
			::SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)pInfo);
		}
		CComQIPtr<IEventProxy> pTangramEvent(EventObj);
		if (pTangramEvent)
		{
			if (pInfo)
			{
				ObjectEventInfo2* pObjectEventInfo2 = nullptr;
				map<LONGLONG, ObjectEventInfo2*>::iterator it2 = pInfo->m_mapEventObj2.find(nIndex);
				if (it2 == pInfo->m_mapEventObj2.end())
				{
					pObjectEventInfo2 = new ObjectEventInfo2();
					pInfo->m_mapEventObj2[nIndex] = pObjectEventInfo2;
				}
				else
					pObjectEventInfo2 = it2->second;
				map<WindowEventType, IEventProxy*>::iterator it = pObjectEventInfo2->m_mapEventObj.find(EventType);
				if (it == pObjectEventInfo2->m_mapEventObj.end())
				{
					pObjectEventInfo2->m_mapEventObj[EventType] = pTangramEvent.p;
					m_pTangramProxy->AttachEvent(pCtrl, EventType);
				}
			}
		}
	}
	return bstrRet;
}

HRESULT CCloudAddinCLRProxy::ActiveCLRMethod(BSTR bstrObjID, BSTR bstrMethod, BSTR bstrParam, BSTR bstrData)
{
	cli::array<Object^, 1>^ pObjs = { BSTR2STRING(bstrParam), BSTR2STRING(bstrData) };
	TangramCLR::Tangram::ActiveMethod(BSTR2STRING(bstrObjID), BSTR2STRING(bstrObjID), pObjs);
	return S_OK;
}

void CCloudAddinCLRProxy::AttachPropertyGridCtrl(HWND hWnd) 
{
	Control^ pCtrl = Control::FromHandle((IntPtr)hWnd);
	if (pCtrl)
	{
		//Control^ parent = pCtrl->Parent;
		//Type^ pType = parent->GetType();
		//String^ strName = pType->Name;
		//String^ name = pType->Assembly->FullName;
		//String^ strNM = pType->Namespace;
		//CString strInfo = GetLibControlsStr(pType->Assembly);
		////parent->GetType()->Assembly->GetFile()
		//cli::array<MethodInfo^>^ pArray = pType->GetMethods();
		////int nCount = pArray->Count;
		//for each(MethodInfo^ pInfo in pArray)
		//{
		//	String^ strname = pInfo->Name;
		//	Debug::Print(strname + "\r\n");
		//}
		PropertyGrid^ pPropertyGrid = (PropertyGrid^)pCtrl;
		m_pVSPropertyGrid = pPropertyGrid;
		m_pVSPropertyGrid->PropertyValueChanged += gcnew PropertyValueChangedEventHandler(&OnPropertyValueChanged);
		m_pVSPropertyGrid->SelectedGridItemChanged += gcnew SelectedGridItemChangedEventHandler(&OnSelectedGridItemChanged);
	}
}

IDispatch* CCloudAddinCLRProxy::CreateCLRObj(BSTR bstrObjID)
{
	Object^ pObj = TangramCLR::Tangram::CreateObject(BSTR2STRING(bstrObjID));
	if (pObj != nullptr)
		return (IDispatch*)Marshal::GetIUnknownForObject(pObj).ToPointer();
	return nullptr;
}

Control^ CCloudAddinCLRProxy::GetCanSelect(Control^ ctrl, bool direct)
{
	int nCount = ctrl->Controls->Count;
	Control^ pCtrl = nullptr;
	if (nCount)
	{
		for (int i = direct ? (nCount - 1):0; direct ? i >= 0 : i <= nCount - 1; direct ? i-- : i++)
		{
			pCtrl = ctrl->Controls[i];
			if (direct && pCtrl->TabStop && pCtrl->Visible && pCtrl->Enabled)
				return pCtrl;
			pCtrl = GetCanSelect(pCtrl, direct);
			if (pCtrl != nullptr)
				return pCtrl;
		}
	}
	else if ((ctrl->CanSelect||ctrl->TabStop)&&ctrl->Visible&&ctrl->Enabled)
		return ctrl;
	return nullptr;
}

HRESULT CCloudAddinCLRProxy::ProcessCtrlMsg(HWND hCtrl,bool bShiftKey)
{
	Control^ pCtrl = (Control^)Control::FromHandle((IntPtr)hCtrl);
	if (pCtrl == nullptr)
		return S_FALSE;
	Control^ pChildCtrl = GetCanSelect(pCtrl, !bShiftKey);
	
	if (pChildCtrl)
		pChildCtrl->Select();
	return S_OK;
}

BOOL CCloudAddinCLRProxy::ProcessFormMsg(HWND hFormWnd, LONG lpMSG, int nMouseButton)
{
	return false;
}

HRESULT CCloudAddinCLRProxy::CloseForms()
{
	FormCollection^ pCollection = Application::OpenForms;
	int nCount = pCollection->Count;
	while (pCollection->Count>0)
	{
		Form^ pForm = pCollection[0];
		pForm->Close();
		delete pForm;
	}
	return S_OK;
}

CString CCloudAddinCLRProxy::GetControlStrs(CString strLib)
{ 
	return GetLibControlsStr(BSTR2STRING(strLib));
}

void CCloudAddinCLRProxy::SelectNode(IWndNode* pNode)
{
	if (m_pPropertyGrid)
	{
		Object^ pObj = nullptr;
		try
		{
			if(pNode)
				pObj = Marshal::GetObjectForIUnknown((IntPtr)pNode);
		}
		catch (...)
		{

		}
		finally
		{
			if (pObj != nullptr)
			{
				try
				{
					m_pPropertyGrid->SelectedObject = pObj;
					//::WndNodeType type;
					//pNode->get_NodeType(&type);
					//if (type == ::WndNodeType::TNT_Splitter)
					//{

					//}
					//else
					//{
					//	int nCount = m_pPropertyGrid->SelectedGridItem->GridItems->Count;
					//	for each(GridItem^ x in m_pPropertyGrid->SelectedGridItem->GridItems)
					//	{
					//		String^ s = x->Value->ToString();
					//		String^ name = x->Label;// ->ToString();
					//	}
					//}
				}
				catch (...)
				{

				}
			}
			else
			{
				m_pPropertyGrid->SelectedObject = nullptr;
			}
		}
		//IWndNode* pParent = NULL;
		//pNode->get_RootNode(&pParent);
		//cli::array<Object^>^ p = { Marshal::GetObjectForIUnknown((IntPtr)pNode), Marshal::GetObjectForIUnknown((IntPtr)pParent) };
		//m_pPropertyGrid->SelectedObject = p;// Marshal::GetObjectForIUnknown((IntPtr)pNode), Marshal::GetObjectForIUnknown((IntPtr)pParent)
	};
}

CString CCloudAddinCLRProxy::GetLibControlsStr(String^ strLib)
{
	Assembly^ m_pDotNetAssembly = nullptr;
	String^ strCtls = L"";
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
	finally
	{
		if (m_pDotNetAssembly != nullptr)
		{
			cli::array<Type^>^ pArray = m_pDotNetAssembly->GetExportedTypes();
			for each (Type^ type in pArray)
			{
				Type^ basetype = type->BaseType;
				while (basetype != nullptr)
				{
					if (basetype == Form::typeid||basetype==Component::typeid)
						break;
					if (basetype == Control::typeid)
					{
						strCtls += type->FullName;
						strCtls += L",";
						break;
					}
					basetype = basetype->BaseType;
				}
			}
		}
	}
	return strCtls;
}

CString CCloudAddinCLRProxy::GetLibControlsStr(Assembly^ Assembly)
{
	String^ strCtls = L"";
	if (Assembly != nullptr)
	{
		cli::array<Type^>^ pArray = Assembly->GetExportedTypes();
		for each (Type^ type in pArray)
		{
			Type^ basetype = type->BaseType;
			while (basetype != nullptr)
			{
				if (basetype == Form::typeid||basetype==Component::typeid)
					break;
				if (basetype == Control::typeid)
				{
					strCtls += type->FullName;
					strCtls += L",";
					break;
				}
				basetype = basetype->BaseType;
			}
		}
	}
	return strCtls;
}

IDispatch* CCloudAddinCLRProxy::TangramCreateObject(BSTR bstrObjID, long hParent, IWndNode* pHostNode)
{
	String^ strID = BSTR2STRING(bstrObjID);
	// Marshal::PtrToStringUni((System::IntPtr)LPTSTR(LPCTSTR(bstrObjID)));
	//if (hParent)
	//	m_pTangramProxy->m_hParentWnd = (HWND)hParent;

	Control^ pObj = static_cast<Control^>(TangramCLR::Tangram::CreateObject(strID));
	if (pObj != nullptr&&pHostNode)
	{
		m_strObjTypeName = pObj->GetType()->Name;
		IWndNode* pRootNode = NULL;
		pHostNode->get_RootNode(&pRootNode);
		m_pTangramProxy->DelegateEvent(pObj, pHostNode);
		HWND hWnd = (HWND)pObj->Handle.ToInt64();
		CComBSTR bstrName(L"");
		pHostNode->get_Name(&bstrName);
		CString strName = OLE2T(bstrName);
		if (strName.CompareNoCase(_T("TangramPropertyGrid")) == 0)
		{
			m_pPropertyGrid = (PropertyGrid^)pObj;
			m_pPropertyGrid->ToolbarVisible = false;
			m_pPropertyGrid->PropertySort = PropertySort::Alphabetical;
		}
		return (IDispatch*)(Marshal::GetIUnknownForObject(pObj).ToInt64());
	}
	return nullptr;
}

BOOL CCloudAddinCLRProxy::IsWinForm(HWND hWnd)
{
	if (hWnd == 0)
		return false;
	IntPtr handle = (IntPtr)hWnd;
	Control^  pForm = Form::FromHandle(handle);
	if (pForm != nullptr)
	{
		Type^ pType = pForm->GetType()->BaseType;
		while (pType != nullptr&&pType->Name != L"Control")
		{
			if (pType->Name == L"Form")
				return true;
			pType = pType->BaseType;
		}
	}
	return false;
}

IDispatch* CCloudAddinCLRProxy::GetCLRControl(IDispatch* CtrlDisp, BSTR bstrNames)
{
	CString strNames = OLE2T(bstrNames);
	if (strNames != _T(""))
	{
		int nPos = strNames.Find(_T(","));
		if (nPos == -1)
		{
			Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)CtrlDisp);
			if (pCtrl != nullptr)
			{
				Control::ControlCollection^ Ctrls = pCtrl->Controls;
				pCtrl = Ctrls[BSTR2STRING(bstrNames)];
				if (pCtrl == nullptr)
				{
					int nIndex = _wtoi(OLE2T(bstrNames));
					pCtrl = Ctrls[nIndex];
				}
				if (pCtrl != nullptr)
					return (IDispatch*)Marshal::GetIDispatchForObject(pCtrl).ToPointer();
			}
			return S_OK;
		}

		Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)CtrlDisp);
		while (nPos != -1)
		{
			CString strName = strNames.Left(nPos);
			if (strName != _T(""))
			{
				if (pCtrl != nullptr)
				{
					Control^ pCtrl2 = pCtrl->Controls[BSTR2STRING(strName)];
					if (pCtrl2 == nullptr)
					{
						if (pCtrl->Controls->Count > 0)
							pCtrl2 = pCtrl->Controls[_wtoi(strName)];
					}
					if (pCtrl2 != nullptr)
						pCtrl = pCtrl2;
					else
						return S_OK;
				}
				else
					return S_OK;
			}
			strNames = strNames.Mid(nPos + 1);
			nPos = strNames.Find(_T(","));
			if (nPos == -1)
			{
				if (pCtrl != nullptr)
				{
					Control^ pCtrl2 = pCtrl->Controls[BSTR2STRING(strNames)];
					if (pCtrl2 == nullptr)
					{
						if (pCtrl->Controls->Count > 0)
							pCtrl2 = pCtrl->Controls[_wtoi(strName)];
					}
					if (pCtrl2 == nullptr)
						return S_OK;
					if (pCtrl2 != nullptr)
						return (IDispatch*)Marshal::GetIDispatchForObject(pCtrl2).ToPointer();
				}
			}
		}
	}
	return NULL;
}

BSTR CCloudAddinCLRProxy::GetCtrlName(IDispatch* _pCtrl)
{
	Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)_pCtrl);
	if (pCtrl != nullptr)
		return STRING2BSTR(pCtrl->Name);
	return L"";
}

HWND CCloudAddinCLRProxy::GetMDIClientHandle(IDispatch* pMDICtrl)
{
	Form^ pCtrl = (Form^)Marshal::GetObjectForIUnknown((IntPtr)pMDICtrl);
	if (pCtrl != nullptr)
	{
		Control^ ctrl = TangramCLR::Tangram::GetMDIClient(pCtrl);
		if (ctrl != nullptr)
			return (HWND)ctrl->Handle.ToInt64();
	}
	return NULL;
}

IDispatch* CCloudAddinCLRProxy::GetCtrlByName(IDispatch* CtrlDisp, BSTR bstrName, bool bFindInChild)
{
	try
	{
		Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)CtrlDisp);
		if (pCtrl != nullptr)
		{
			cli::array<Control^, 1>^ pArray = pCtrl->Controls->Find(BSTR2STRING(bstrName), bFindInChild);
			if (pArray != nullptr&&pArray->Length)
				return (IDispatch*)Marshal::GetIDispatchForObject(pArray[0]).ToPointer();
		}

	}
	catch (System::Exception^)
	{

	}
	return NULL;
}

void CCloudAddinCLRProxy::InitComponent(Control^ pCtrl)
{
	if (pCtrl != nullptr)
	{
		Control::ControlCollection^ pCtrls = pCtrl->Controls;
		pCtrl->ControlAdded += gcnew System::Windows::Forms::ControlEventHandler(&OnControlAdded);
		pCtrl->ControlRemoved += gcnew System::Windows::Forms::ControlEventHandler(&OnControlRemoved);
		int nCount = pCtrls->Count;
		if (nCount)
		{
			for each(Control^ ctrl in pCtrls)
			{
				CString strCtrlName = OLE2T(STRING2BSTR(ctrl->Name));
				IntPtr nHandle = ctrl->Handle;
				LONGLONG hWnd = nHandle.ToInt64();
				HWND _hWnd = (HWND)hWnd;
				if (::IsWindow(_hWnd))
				{
					if (ctrl->Dock == DockStyle::Fill)
					{
						::SendMessage(m_hCurDesignObj, WM_TANGRAMMSG, (WPARAM)hWnd, (LPARAM)strCtrlName.GetBuffer());
					}
					InitComponent(ctrl);
				}
			}
		}
	}
}

HWND CCloudAddinCLRProxy::OnInitialDesigner(long hWnd, CString& strName)
{
	HWND hRet = NULL;
	try
	{
		IntPtr handle = (IntPtr)hWnd;
		m_hCurDesignObj = (HWND)hWnd;
		Control^  pCtrl = Control::FromHandle(handle);
		if (pCtrl != nullptr)
		{
			strName = pCtrl->Name;
			Control^ MdiClientCtrl = nullptr;
			Type^ pType = pCtrl->GetType();
			Type^ pFormType = Form::typeid;
			while (pType != pFormType)
			{
				pType = pType->BaseType;
				if (pType == Control::typeid)
					break;
			}
			if (pType == pFormType)
			{
				MdiClientCtrl = TangramCLR::Tangram::GetMDIClient((Form^)pCtrl);
				if (MdiClientCtrl != nullptr)
					return (HWND)MdiClientCtrl->Handle.ToInt64();
			}
			InitComponent(pCtrl);
		}
	}
	catch (System::Exception^)
	{

	}
	return hRet;
}

CString CCloudAddinCLRProxy::GetCtrlInfo(long hWnd)
{	
	CString strInfo = _T("");
	try
	{
		IntPtr handle = (IntPtr)hWnd;
		Control^  pCtrl = Control::FromHandle(handle);
		if (pCtrl != nullptr)
		{
			Control^ pTopParent = pCtrl->Parent;
			bool bTop = false;
			if (pTopParent == nullptr)
			{
				pTopParent=pCtrl;
				bTop = true;
			}
			while (bTop == false)
			{
				Control^ _pCtrl = pTopParent->Parent;
				if (_pCtrl != nullptr)
					pTopParent = _pCtrl;
				else
				{
					bTop = true;
				}
			}
			HWND hTop = (HWND)pTopParent->Handle.ToInt64();
			TCHAR	m_szBuffer[MAX_PATH];
			::GetClassName(hTop, m_szBuffer, MAX_PATH);
			CString strClassName = CString(m_szBuffer);
			if (strClassName.CompareNoCase(_T("Tangram Window Class")) == 0)
			{
				return strInfo;
			}
			hTop = ::GetParent(hTop);
			CString strText = _T("");
			::GetWindowText(hTop, m_szBuffer, MAX_PATH);
			strText = CString(m_szBuffer);
			if (strText.CompareNoCase(_T("OverlayControl"))==0)
				return strInfo;
			String^ name = pCtrl->Name;
			strInfo = OLE2T(STRING2BSTR(name));
			if (pCtrl->Dock == DockStyle::Fill)
			{
				strInfo += _T("|1");
			}
			else
			{
				Type^ pType = pCtrl->GetType();
				if(pType->FullName==L"System.Windows.Forms.TabPage"|| pType->FullName == L"System.Windows.Forms.SplitterPanel")
					strInfo += _T("|1");
				else
					strInfo += _T("|0");
			}
		}		
	}
	catch (System::Exception^)
	{

	}
	return strInfo;
}

BSTR CCloudAddinCLRProxy::GetCtrlValueByName(IDispatch* CtrlDisp, BSTR bstrName, bool bFindInChild)
{
	try
	{
		Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)CtrlDisp);
		if (pCtrl != nullptr)
		{
			cli::array<Control^, 1>^ pArray = pCtrl->Controls->Find(BSTR2STRING(bstrName), bFindInChild);
			if (pArray != nullptr&&pArray->Length)
			{
				return STRING2BSTR(pArray[0]->Text);
			}
		}
	}
	catch (System::Exception^)
	{

	}
	return NULL;
}

HWND CCloudAddinCLRProxy::GetCtrlHandle(IDispatch* _pCtrl)
{
	Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)_pCtrl);
	if (pCtrl != nullptr)
		return (HWND)pCtrl->Handle.ToInt64();
	return 0;
}

HWND CCloudAddinCLRProxy::IsCtrlCanNavigate(IDispatch* ctrl)
{
	Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)ctrl);
	if (pCtrl != nullptr)
	{
		if (pCtrl->Dock == DockStyle::Fill)
			return (HWND)pCtrl->Handle.ToInt64();
	}
	return 0;
}

void CCloudAddinCLRProxy::TangramAction(BSTR bstrXml, IWndNode* pNode)
{
	CString strXml = OLE2T(bstrXml);
	if (strXml != _T(""))
	{
		CTangramXmlParse m_Parse;
		WndNode^ pWindowNode = (WndNode^)theAppProxy._createObject<IWndNode, WndNode>(pNode);
		if (pWindowNode&&m_Parse.LoadXml(strXml))
		{
			int nType = m_Parse.attrInt(_T("Type"), 0);
			switch (nType)
			{
			case 5:
				if (pWindowNode != nullptr)
				{
				}
				break;
			default:
				{
					CString strID = m_Parse.attr(_T("ObjID"), _T(""));
					CString strMethod = m_Parse.attr(_T("Method"), _T(""));
					if (strID != _T("") && strMethod != _T(""))
					{
						cli::array<Object^, 1>^ pObjs = { BSTR2STRING(strXml), pWindowNode };
						TangramCLR::Tangram::ActiveMethod(BSTR2STRING(strID), BSTR2STRING(strMethod), pObjs);
					}
				}
				break;
			}
		}
	}
}


bool CCloudAddinCLRProxy::_insertObject(Object^ key, Object^ val)
{
	Hashtable^ htObjects = (Hashtable^)m_htObjects;
	htObjects[key] = val;
	return true;
}

Object^ CCloudAddinCLRProxy::_getObject(Object^ key)
{
	Hashtable^ htObjects = (Hashtable^)m_htObjects;
	return htObjects[key];
}

bool CCloudAddinCLRProxy::_removeObject(Object^ key)
{
	Hashtable^ htObjects = (Hashtable^)m_htObjects;

	if (htObjects->ContainsKey(key))
	{
		htObjects->Remove(key);
		return true;
	}
	return false;
}

void CTangramNodeEvent::OnExtendComplete()
{
	if (m_pTangramNodeCLREvent)
		m_pTangramNodeCLREvent->OnExtendComplete(NULL);
}

void CTangramNodeEvent::OnDestroy()
{
	if (m_pTangramNodeCLREvent)
	{
		m_pTangramNodeCLREvent->OnDestroy();
		delete m_pTangramNodeCLREvent;
	}
}

void CTangramNodeEvent::OnDocumentComplete(IDispatch* pDocdisp, BSTR bstrUrl)
{
	if (m_pWndNode!=nullptr)
		m_pTangramNodeCLREvent->OnDocumentComplete(pDocdisp, bstrUrl);
}

void CTangramNodeEvent::OnNodeAddInCreated(IDispatch* pAddIndisp, BSTR bstrAddInID, BSTR bstrAddInXml)
{
	if (m_pWndNode != nullptr)
		m_pTangramNodeCLREvent->OnNodeAddInCreated(pAddIndisp, bstrAddInID, bstrAddInXml);
}

void CTangramNodeEvent::OnNodeAddInsCreated()
{
	if (m_pWndNode != nullptr)
		m_pTangramNodeCLREvent->OnNodeAddInsCreated();
}

void CCloudAddinCLRApp::OnClose()
{
	AtlTrace(_T("*************CCloudAddinCLRApp::OnClose:  ****************\n"));
	HRESULT hr = DispEventUnadvise(m_pTangram);
	if (m_pCloudAddinCLREvent)
		delete m_pCloudAddinCLREvent;
}

void CCloudAddinCLRApp::OnExtendComplete(long hWnd, BSTR bstrUrl, IWndNode* pRootNode)
{
	if (m_pCloudAddinCLREvent)
		m_pCloudAddinCLREvent->OnExtendComplete(hWnd,bstrUrl,pRootNode);
}

CWndPageClrEvent::CWndPageClrEvent()
{

}

CWndPageClrEvent::~CWndPageClrEvent()
{
}

void __stdcall  CWndPageClrEvent::OnDestroy()
{
	if (m_pPage)
		delete m_pPage;
}

void CWndPageClrEvent::OnInitialize(IDispatch* pHtmlWnd, BSTR bstrUrl)
{
	Object^ pObj = m_pPage;
	TangramCLR::WndPage^ pPage = static_cast<TangramCLR::WndPage^>(pObj);
	pPage->Fire_OnDocumentComplete(pPage, Marshal::GetObjectForIUnknown((IntPtr)pHtmlWnd), BSTR2STRING(bstrUrl));
}


void CCloudAddinCLRProxy::OnControlAdded(Object ^sender, ControlEventArgs ^e)
{
	Control^ ctrl = e->Control;
	String^ strName = ctrl->Name;
	ctrl->DockChanged += gcnew EventHandler(&OnDockChanged);
	ctrl->ParentChanged += gcnew EventHandler(&OnParentChanged);
	ctrl->ControlAdded += gcnew ControlEventHandler(&OnControlAdded);
	ctrl->ControlRemoved += gcnew ControlEventHandler(&OnControlRemoved);
}


Control^ CCloudAddinCLRProxy::GetParentByName(Control^ ctrl, String^ name)
{
	Control^ parent = ctrl->Parent;
	String^ caption = parent->Text;
	while (caption != name)
	{
		Control^ pOld = parent;
		parent = parent->Parent;
		caption = parent->Text;
		if (caption == name)
		{
			parent = pOld;
			return parent;
		}
	}
	return nullptr;
}

void CCloudAddinCLRProxy::OnDockChanged(Object ^sender, EventArgs ^e)
{
	Control^ ctrl = (Control^)sender;
	DockStyle m_style = ctrl->Dock;
	if (m_style == DockStyle::Fill)
	{
		Control^ parent = GetParentByName(ctrl, L"OverlayControl");
		HWND hWnd = (HWND)parent->Handle.ToInt64();
		CString strName = ctrl->Name;
		HWND hCtrl = (HWND)ctrl->Handle.ToInt64();
		theAppProxy.m_mapDesignObj[hCtrl] = hWnd;
		::SendMessage(hWnd, WM_TANGRAMMSG, (WPARAM)hCtrl, (LPARAM)strName.GetBuffer());
	}
}

void CCloudAddinCLRProxy::OnControlRemoved(Object ^sender, ControlEventArgs ^e)
{
	Control^ ctrl = e->Control;
	String^ strName = ctrl->Name;
	HWND hWnd = (HWND)ctrl->Handle.ToInt64();
	map<HWND, HWND>::iterator it = theAppProxy.m_mapDesignObj.find(hWnd);
	if(it!=theAppProxy.m_mapDesignObj.end())
		::SendMessage(it->second, WM_TANGRAMMSG, (WPARAM)ctrl->Handle.ToInt64(), 1963);
}

void CCloudAddinCLRProxy::OnParentChanged(Object ^sender, EventArgs ^e)
{
}

void CCloudAddinCLRProxy::OnSelectedGridItemChanged(Object ^sender, SelectedGridItemChangedEventArgs ^e)
{
	e->NewSelection->PropertyDescriptor;
}

void CCloudAddinCLRProxy::OnPropertyValueChanged(Object ^s, PropertyValueChangedEventArgs ^e)
{
	String^ strVal = e->ChangedItem->Value->ToString();
	Object^ pObj = theAppProxy.m_pVSPropertyGrid->SelectedObject;
	IDispatch* pDisp = (IDispatch*)Marshal::GetIUnknownForObject(pObj).ToPointer();
	CComQIPtr<IWndNode> pNode(pDisp);
	if (pNode)
	{

	}
	else
	{
		Control^ pCtrl = (Control^)pObj;
		if (pCtrl)
		{
			HWND hwnd = (HWND)pCtrl->Handle.ToInt64();
			String^ strName = e->ChangedItem->Label;
			if (strName == L"(Name)")
			{
				String^ strVal = e->ChangedItem->Value->ToString();
				CString strNewName = STRING2BSTR(strVal);
				::SendMessage(hwnd, WM_NAMECHANGED, (WPARAM)strNewName.GetBuffer(), 0);
			}
		}
	}
}
