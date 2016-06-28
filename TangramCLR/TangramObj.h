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

#pragma once
#include "Tangram.h"
#include "TangramClrProxy.h"
#include <map>

using namespace std;
using namespace System;
using namespace System::Windows;
using namespace System::Reflection;
using namespace System::Windows::Forms;
using namespace System::Runtime::InteropServices;
using namespace System::ComponentModel;
using namespace System::Collections;
using namespace System::Collections::Generic;
using namespace System::IO;
using namespace System::Threading;
using namespace System::Threading::Tasks;

using System::Runtime::InteropServices::Marshal;

extern CCloudAddinCLRProxy theAppProxy;;
class CTangramNodeEvent;


namespace TangramCLR 
{
	//[Guid("54499D5E-AC2F-4F8B-9782-C00A9BB2F4E2")]
	//[ComVisibleAttribute(true)]
	//[InterfaceTypeAttribute(ComInterfaceType::InterfaceIsIDispatch)]
	//public interface class IManagerExtender
	//{
	//	[DispId(0x000001)]
	//	virtual void CloseForm(System::Object^ pForm);
	//	[DispId(0x000002)]
	//	virtual void OnCloseForm(long long hFormWnd);
	//};

	/// <summary>
	/// Summary for Tangram
	/// </summary>
	ref class WndNode;
	public enum class WndNodeType
	{
		TNT_Blank			= 0x00000001,
		TNT_ActiveX			= 0x00000002,
		TNT_Splitter		= 0x00000004,
		TNT_Tabbed			= 0x00000008,

		TNT_CLR_Control		= 0x00000010,
		TNT_CLR_Form		= 0x00000020,
		TNT_CLR_Window		= 0x00000040,

		TNT_View			= 0x00000080		
	};

	template<typename T> 
	private ref class TangramBaseEnumerator : public System::Collections::IEnumerator
	{
	public:
		TangramBaseEnumerator(T^ v, int iCount)
		{
			m_pT = v;
			nCount = iCount;
			Reset();
		}

		virtual property Object^ Current
		{
			Object^ get()
			{
				if(index < 0 || index >= nCount)
				{
					throw gcnew InvalidOperationException();
				}
				return m_pT->default[index];
			}
		}

		virtual bool MoveNext()
		{
			index++;
			return (index < nCount);
		} 

		virtual void Reset()
		{
			index = -1;
		}

	private:
		int index;		
		long nCount;
		T^ m_pT;
	};

	ref class WndPage;
	ref class WndNode;
	
	public ref class WndNodeCollection : public System::Collections::IEnumerable
	{
	public:
		WndNodeCollection(IWndNodeCollection* pNodes)
		{		
			SetNewNodeCollection(pNodes);
		};

		~WndNodeCollection()
		{
			SetNewNodeCollection(NULL);
		}
		
	private:
		IWndNodeCollection* m_pNodeCollection;

		void SetNewNodeCollection(IWndNodeCollection* pNodes)
		{
			if (m_pNodeCollection != NULL)
			{
				m_pNodeCollection->Release();
				m_pNodeCollection = NULL;
			}
			if (pNodes != NULL)
			{
				m_pNodeCollection = pNodes;
				m_pNodeCollection->AddRef();					
			}
		}
	public:
		virtual System::Collections::IEnumerator^ GetEnumerator()
		{
			return gcnew TangramBaseEnumerator<WndNodeCollection>(this,NodeCount);
		}
		
		property WndNode^ default[int]
		{
			WndNode^ get(int iIndex);
		}

		property int NodeCount
		{
			int get()
			{
				long n = 0;
				m_pNodeCollection->get_NodeCount(&n);
				return n;
			}
		}		
	};

	ref class WndFrame;
	public ref class WndNode
	{	
	public:
		WndNode(IWndNode* pNode);
		~WndNode();

		IWndNode* m_pWndNode;
		CTangramNodeEvent* m_pTangramNodeEvent;
		//virtual void CloseForm(Object^ pForm){};
		//virtual void OnCloseForm(long long hFormWnd){};

		delegate void NodeAddInCreated(WndNode^ sender, Object^ pAddIndisp, String^ bstrAddInID, String^ bstrAddInXml);
		event NodeAddInCreated^ OnNodeAddInCreated;
		void Fire_NodeAddInCreated(WndNode^ sender, Object^ pAddIndisp, String^ bstrAddInID, String^ bstrAddInXml)
		{
			OnNodeAddInCreated(sender, pAddIndisp, bstrAddInID, bstrAddInXml);
		}

		delegate void NodeAddInsCreated(WndNode^ sender);
		event NodeAddInsCreated^ OnNodeAddInsCreated;
		void Fire_NodeAddInsCreated(WndNode^ sender)
		{
			OnNodeAddInsCreated(sender);
		}

		delegate void ExtendComplete(WndNode^ sender);
		event ExtendComplete^ OnExtendComplete;
		void Fire_ExtendComplete(WndNode^ sender)
		{
			OnExtendComplete(sender);
		}

		delegate void Destroy(WndNode^ sender);
		event Destroy^ OnDestroy;
		void Fire_OnDestroy(WndNode^ sender)
		{
			OnDestroy(sender);
		}

		delegate void DocumentComplete(WndNode^ sender, Object^ pHtmlDoc, String^ strURL);
		event DocumentComplete^ OnDocumentComplete;
		void Fire_OnDocumentComplete(WndNode^ sender, Object^ pHtmlDoc, String^ strURL)
		{
			OnDocumentComplete(sender, pHtmlDoc, strURL);
		}

	private:
		HWND m_hWnd;
		Object^ m_pHostObj = nullptr;
		Dictionary<String^, MethodInfo^>^	m_pTangramCLRMethodDic = nullptr;
		Dictionary<String^, Object^>^	m_pTangramPlugInDic = nullptr;
		void SetNewNode(IWndNode* pNode)
		{
			if (m_pWndNode != NULL)
			{
				m_pWndNode = NULL;
			}

			if (pNode != NULL)
			{
				m_pWndNode = pNode;
			}
		}

	public:		
		void Init();
		//void SaveTangramDoc(String^ m_strName);
		Object^ ActiveMethod(String^ strMethod, cli::array<Object^, 1>^ p);
		//[DispId(0x000001)]
		//[ComVisibleAttribute(true)]
		//Object^ ExecScript(String^ strScript);
		WndNode^ Extend(String^ strKey, String^ strXml);

		property String^ Caption 
		{
			String^ get();
			void set(String^ strCaption)
			{
				m_pWndNode->put_Caption(STRING2BSTR(strCaption));
			}
		}

		property WndPage^ WndPage
		{
			TangramCLR::WndPage^ get();
		}

		property String^ Name
		{
			String^ get()
			{
				if(m_pWndNode)
				{
					CComBSTR bstrCap("");
					m_pWndNode->get_Name(&bstrCap);
					return BSTR2STRING(bstrCap);
				}
				return "";
			}
		}

		property Object^ XObject
		{
			Object^ get()
			{
				if (m_pHostObj != nullptr)
					return m_pHostObj;

				VARIANT var;
				m_pWndNode->get_XObject(&var);
				
				try
				{
					m_pHostObj = Marshal::GetObjectForNativeVariant((System::IntPtr)&var);
				}
				catch (InvalidOleVariantTypeException^ e)
				{
					e->Message;
				}
				catch (...)
				{

				}
				return m_pHostObj;
			}
		}

		property Object^ Extender
		{
			Object^ get()
			{
				Object^ pRetObject = nullptr;
				IDispatch* pDisp = NULL;
				m_pWndNode->get_Extender(&pDisp);
				if (pDisp)
					pRetObject = Marshal::GetObjectForIUnknown((IntPtr)pDisp); 
				return pRetObject;
			}
			void set(Object^ obj)
			{
				IDispatch* pDisp = (IDispatch*)Marshal::GetIUnknownForObject(obj).ToPointer();
				m_pWndNode->put_Extender(pDisp);
			}
		}

		property Object^ Tag
		{
			Object^ get()
			{
				Object^ pRetObject = nullptr;
				VARIANT var;
				m_pWndNode->get_Tag(&var);	
				try
				{
					pRetObject = Marshal::GetObjectForNativeVariant((System::IntPtr)&var);
				}
				catch (InvalidOleVariantTypeException^ e)
				{

				}
				catch (NotSupportedException^ e)
				{

				}
				return pRetObject;
			}

			void set(Object^ obj)
			{
				try
				{
					VARIANT var;
					Marshal::GetNativeVariantForObject(obj,(System::IntPtr)&var);
					m_pWndNode->put_Tag(var);
				}
				catch (ArgumentException^ e)
				{
					e->Data->ToString();
				}
			}
		}

		property WndFrame^ HostFrame
		{
			WndFrame^ get()
			{
				CComPtr<IWndFrame> pTangramFrame;
				m_pWndNode->get_Frame(&pTangramFrame);

				WndFrame^ pFrame = theAppProxy._createObject<IWndFrame, WndFrame>(pTangramFrame);
				return pFrame;
			}
		}

		property Object^ PlugIn[String^]
		{
			Object^ get(String^ strName);
		}

		property WndNodeCollection^ ChildNodes
		{
			WndNodeCollection^ get();
		}

		property WndNodeCollection^ Objects[WndNodeType]
		{
			WndNodeCollection^ get(WndNodeType nValue)
			{
				CComPtr<IWndNodeCollection> pNodes = NULL;
				m_pWndNode->get_Objects((long)nValue,&pNodes);
				return gcnew WndNodeCollection(pNodes);				
			}
		}

		property WndNode^ RootNode
		{
			WndNode^ get()
			{
				CComPtr<IWndNode> pRootNode;
				m_pWndNode->get_RootNode(&pRootNode);

				return theAppProxy._createObject<IWndNode, WndNode>(pRootNode);
			}
		}

		property WndNode^ ParentNode
		{
			WndNode^ get()
			{
				CComPtr<IWndNode> pRootNode;
				m_pWndNode->get_ParentNode(&pRootNode);

				return theAppProxy._createObject<IWndNode, WndNode>(pRootNode);
			}
		}

		property int Col
		{
			int get()
			{
				long nValue = 0;
				m_pWndNode->get_Col(&nValue);
				return nValue;
			}
		}

		property int Row
		{
			int get()
			{
				long nValue = 0;
				m_pWndNode->get_Row(&nValue);
				return nValue;
			}
		}

		property int Rows
		{
			int get()
			{
				long nValue = 0;
				m_pWndNode->get_Rows(&nValue);
				return nValue;
			}
		}

		property int Cols
		{
			int get()
			{
				long nValue = 0;
				m_pWndNode->get_Cols(&nValue);
				return nValue;
			}
		}

		property WndNodeType NodeType
		{
			WndNodeType get()
			{
				::WndNodeType type;
				m_pWndNode->get_NodeType(&type);
				return (TangramCLR::WndNodeType)type;
			}
		}

		property long Handle
		{
			long get()
			{
				if (m_hWnd)
					return (long)m_hWnd;
				long h = 0;
				m_pWndNode->get_Handle(&h);
				m_hWnd = (HWND)h;
				return h;
			}
		}

		property String^ Attribute[String^]
		{
			String^ get(String^ strKey)
			{
				BSTR bstrVal;
				m_pWndNode->get_Attribute(STRING2BSTR(strKey),&bstrVal);

				return BSTR2STRING(bstrVal);
			}

			void set(String^ strKey, String^ strVal)
			{
				m_pWndNode->put_Attribute(
					STRING2BSTR(strKey),
					STRING2BSTR(strVal));
			}
		}

public:
	WndNode^ GetNode(int nRow, int nCol)
	{
		IWndNode* pNode;		
		m_pWndNode->GetNode(nRow,nCol,&pNode);

		WndNode^ tNode = theAppProxy._createObject<IWndNode, WndNode>(pNode);

		return tNode;
	}

	int GetNodes(String^ strName, WndNode^% pNode, WndNodeCollection^% pNodeCollection)
	{
		IWndNodeCollection* pNodes;
		IWndNode* pTNode;

		pNode = nullptr;
		pNodeCollection = nullptr;

		long nCount;
		m_pWndNode->GetNodes(STRING2BSTR(strName),
			&pTNode,&pNodes,&nCount);

		pNode = theAppProxy._createObject<IWndNode, WndNode>(pTNode);

		pNodeCollection = gcnew WndNodeCollection(pNodes);
		
		pNodes->Release();

		return nCount;
	}

	//void Unload()
	//{
	//	m_pWndNode->Unload();
	//}

	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>

		WndNodeCollection^ m_pChildNodes;
	};

	public ref class WndFrame
	{
	public:
		WndFrame(IWndFrame* pFrame)
		{
			m_pWndFrame = pFrame;
		}
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~WndFrame()
		{
		}

	protected:
		IWndFrame* m_pWndFrame;

	public:
		WndNode^ ExtendXml(String^ bstrXml, String^ bstrKeyName);
		void Attach(bool bAttach)
		{
			if(m_pWndFrame)
			{
				if(bAttach)
				{
					m_pWndFrame->Attach();
				}
				else
				{
					m_pWndFrame->Detach();
				}
			}
		}

		property Object^ FrameData[String^]
		{
			Object^ get(String^ iIndex);
			void set(String^ iIndex, Object^ newVal);
		}
			
		property WndNodeCollection^ RootNodes
		{
			WndNodeCollection^ get()
			{
				CComPtr<IWndNodeCollection> pNodes = NULL;
				m_pWndFrame->get_RootNodes(&pNodes);

				return gcnew WndNodeCollection(pNodes.p);
			}
		}
	};

	//[ComSourceInterfacesAttribute(TangramCLR::IManagerExtender::typeid)]
	public ref class Tangram
	{
	public:
		Tangram();
		~Tangram();
	private:
		IntPtr										m_nMDIClientHandle;
		static Tangram	^							m_pManager;
		static Dictionary<String^, Type^>^			m_pTangramCLRTypeDic = gcnew Dictionary<String^, Type^>();
		static Dictionary<String^, MethodInfo^>^	m_pTangramCLRMethodDic = gcnew Dictionary<String^, MethodInfo^>();

	public:
		static Dictionary<String^, Object^>^		m_pTangramCLRObjDic = gcnew Dictionary<String^, Object^>();
		static Func<Object^, String^>^ m_pTangramInitGlobalTask;
		static String^ ComputeHash(String^ source);
		static String^ TangramGlobalFunc(Object^);
		static AutoResetEvent^	m_pAppQuitEvent = nullptr;
		static AutoResetEvent^	m_pPlatformStartupCompletedEvent = nullptr;
		static WndPage^ CreateWndPage(Control^ ctrl,Object^ ExternalObj);
		static Object^ CreateObject(String^ ObjID);
		static Object^ ActiveMethod(String^ strObjID, String^ strMethod, cli::array<Object^, 1>^ p);
		static Object^ ActiveObjectMethod(Object^ pObj, String^ strMethod, cli::array<Object^, 1>^ p);
		static Control^ GetMDIClient(Form^ pForm);
		//static WndNode^ ExtendCtrl(Control^ pCtrl, String^ strFile);		
		static int InitTangramCLR(String^ pwzArgument);

		static Tangram^ GetTangram()
		{
			 if(m_pManager==nullptr)
				 m_pManager = gcnew Tangram();
			 return m_pManager;
		}
		
		static Dictionary<String^, WndNode^>^ InitObj(Control^ pObj, String^ strXml);

		static String^ GetObjAssemblyName(Object^ obj)
		{
			 if(obj==nullptr)
				 return L"";
			 System::Reflection::Assembly^ a = nullptr;
			 String^ strName = L"";
			 try
			 {
				a = System::Reflection::Assembly::GetAssembly(obj->GetType());
			 }
			 catch (System::ArgumentNullException^)
			 {

			 }
			 finally
			 {
				 if (a != nullptr)
				 {
					 strName = a->FullName;
					 strName = strName->Substring(0, strName->IndexOf(","));
				 }
			 }
			 return strName;
		}

		static String^ GetTangramRes(Object^ obj, String^ resName)
		{
			 if(obj==nullptr)
				 return L"";
			 if (resName == L"" || resName == nullptr)
				 resName = L"tangramresource.xml";
			 System::Reflection::Assembly^ a = nullptr;
			 String^ strName = L"";
			 try
			 {
				a = System::Reflection::Assembly::GetAssembly(obj->GetType());
			 }
			 catch (System::ArgumentNullException^)
			 {

			 }
			 finally
			 {
				 if (a != nullptr)
				 {
					 strName = a->FullName;
					 strName = strName->Substring(0, strName->IndexOf(","));
				 }
			 }

			 System::IO::Stream^ sm = nullptr;
			 try
			 {
				 sm = a->GetManifestResourceStream(strName + "." + resName);
			 }
			 catch (...)
			 {

			 }
			 finally
			 {
				 if (sm != nullptr)
				 {
					 cli::array<byte,1>^ bs = gcnew cli::array<byte,1>(sm->Length);
					 sm->Read(bs, 0, (int)sm->Length);
					 sm->Close();

					 System::Text::UTF8Encoding^ con = gcnew System::Text::UTF8Encoding();

					 strName = con->GetString(bs);
				 }
			 }
			 return strName;
		}

		static property Object^ Application 
		{
			Object^ get();
			void set(Object^ obj);
		}

		static property WndNode^ CreatingNode 
		{
			WndNode^ get();
		}

		delegate void Close();
		event Close^ OnClose;
		void Fire_OnClose();

		delegate void ExtendComplete(IntPtr hWnd, String^ bstrUrl, WndNode^ pRootNode);
		event ExtendComplete^ OnExtendComplete;
		void Fire_OnExtendComplete(IntPtr hWnd, String^ bstrUrl, WndNode^ pRootNode)
		{
			OnExtendComplete(hWnd, bstrUrl, pRootNode);
		}

		//delegate void DocumentComplete(Object^ Doc, String^ bstrUrl);
		//event DocumentComplete^ OnDocumentComplete;
		//void Fire_OnDocumentComplete(Object^ Doc, String^ bstrUrl)
		//{
		//	OnDocumentComplete(Doc, bstrUrl);
		//}

		property String^ AppKeyValue[String^]
		{
			String^ get(String^ iIndex);
			void set(String^ iIndex, String^ newVal);
		}
	};

	/// <summary>
	/// WndPage 
	/// </summary>
	//[ComSourceInterfacesAttribute(TangramCLR::IManagerExtender::typeid)]
	public ref class WndPage : public IWin32Window
	{
	public:
		WndPage(void);
		WndPage(IWndPage* pPage);

		property IntPtr Handle
		{
			virtual IntPtr get()
			{
				long h = 0;
				m_pPage->get_Handle(&h);
				return (IntPtr)h;
			};
		}

		property String^ URL
		{
			String^ get()
			{
				String^ strRet = nullptr;
				CComBSTR bstr(L"");
				m_pPage->get_URL(&bstr);
				strRet = BSTR2STRING(bstr); 
				return strRet;
			}
			void set(String^ newVal)
			{
 				CComBSTR bstr = STRING2BSTR(newVal);
				m_pPage->put_URL(bstr);
			}
		}

		property Object^ External
		{
			void set(Object^ newVal)
			{
				if (newVal != nullptr)
				{
					IDispatch* pDisp = (IDispatch*)Marshal::GetIUnknownForObject(newVal).ToPointer();
					if (pDisp)
						m_pPage->put_External(pDisp);
				}
			}
		}

		property String^ FrameName[Control^]
		{
			String^ get(Control^ ctrl)
			{
				if (ctrl != nullptr)
				{
					long hWnd = ctrl->Handle.ToInt64();
					BSTR bstrName = ::SysAllocString(L"");
					m_pPage->get_FrameName(hWnd, &bstrName);
					String^ strRet = BSTR2STRING(bstrName);
					::SysFreeString(bstrName);
					return strRet;
				}
				return String::Empty;
			}
		}

		property String^ FrameName[IntPtr]
		{
			String^ get(IntPtr handle)
			{
				if (::IsWindow((HWND)handle.ToInt64()))
				{
					BSTR bstrName = ::SysAllocString(L"");
					m_pPage->get_FrameName(handle.ToInt32(), &bstrName);
					String^ strRet = BSTR2STRING(bstrName);
					::SysFreeString(bstrName);
					return strRet;
				}
				return String::Empty;
			}
		}

		property Object^ Extender[String^]
		{
			Object^ get(String^ strName)
			{
				BSTR bstrName = STRING2BSTR(strName);
				CComPtr<IDispatch> pDisp;
				m_pPage->get_Extender(bstrName, &pDisp);
				::SysFreeString(bstrName);
				return Marshal::GetObjectForIUnknown((IntPtr)pDisp.p);
			}

			void set(String^ strName, Object^ newObj)
			{
				IDispatch* pDisp = (IDispatch*)Marshal::GetIDispatchForObject(newObj).ToPointer();
				m_pPage->put_Extender(STRING2BSTR(strName), pDisp);
			}
		}

		property WndFrame^ Frames[Object^]
		{
			WndFrame^ get(Object^ obj)
			{
				CComVariant m_v;
				IntPtr handle = (IntPtr)&m_v;
				Marshal::GetNativeVariantForObject(obj, handle);
				CComPtr<IWndFrame> pFrame;
				m_pPage->get_Frame(m_v, &pFrame);
				if (pFrame)
				{
					return theAppProxy._createObject<IWndFrame, WndFrame>(pFrame);
				}
				return nullptr;
			}
		}

		WndNode^ GetWndNode(String^ bstrXml, String^ bstrFrameName);

		void AddObjToHtml(String^ strObjName, bool bConnectEvent, Object^ pObj);

		WndFrame^ CreateFrame(Control^ ctrl, String^ strName)
		{
			if (ctrl != nullptr&&String::IsNullOrEmpty(strName) == false)
			{
				CComPtr<IWndFrame> pFrame;
				m_pPage->CreateFrame(CComVariant(0), CComVariant(ctrl->Handle.ToInt64()), STRING2BSTR(strName), &pFrame);
				if (pFrame)
				{
					return theAppProxy._createObject<IWndFrame, WndFrame>(pFrame);
				}
			}
			return nullptr;
		}

		WndFrame^ CreateFrame(IntPtr handle,String^ strName)
		{
			if (::IsWindow((HWND)handle.ToInt64()) && String::IsNullOrEmpty(strName) == false)
			{
				CComPtr<IWndFrame> pFrame;
				m_pPage->CreateFrame(CComVariant(0), CComVariant(handle.ToInt64()), STRING2BSTR(strName), &pFrame);
				if (pFrame)
				{
					return theAppProxy._createObject<IWndFrame, WndFrame>(pFrame);
				}
			}
			return nullptr;
		}

		delegate void DocumentComplete(WndPage^ sender, Object^ pHtmlDoc, String^ strURL);
		event DocumentComplete^ OnDocumentComplete;
		void Fire_OnDocumentComplete(WndPage^ sender, Object^ pHtmlDoc, String^ strURL)
		{
			OnDocumentComplete(sender, pHtmlDoc, strURL);
		}
		delegate void Destroy(WndNode^ sender);
		event Destroy^ OnDestroy;
		void Fire_OnDestroy(WndNode^ sender)
		{
			OnDestroy(sender);
		}
	protected:
		IWndPage*			m_pPage;
		CWndPageClrEvent*	m_pTangramClrEvent;

		~WndPage();
	}; 
}
