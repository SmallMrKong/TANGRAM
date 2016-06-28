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

#include "StdAfx.h"
#include "TangramObj.h"
#include "TangramProxy.h"

namespace TangramCLR
{
	TangramProxy::TangramProxy(void)
	{
		m_pCurrentNode = nullptr;
		m_CtrlEventDic = nullptr;
		theAppProxy.m_pTangramProxy = this;
	}

	void TangramProxy::AttachHandleDestroyedEvent(Control^ pCtrl)
	{
		pCtrl->HandleDestroyed += gcnew EventHandler(theAppProxy.m_pTangramProxy, &TangramProxy::OnHandleDestroyed);
	}

	void TangramProxy::AttachEvent(Control^ pCtrl, WindowEventType EventType)
	{
		if (m_CtrlEventDic == nullptr)
			m_CtrlEventDic = gcnew Dictionary < Control^, String^ > ;
		String^ strVal = L"";
		int nPos = -1;
		CString str = _T("");
		str.Format(_T(",%d,"),EventType);
		String^ _strKey = BSTR2STRING(str);
		if (m_CtrlEventDic->TryGetValue(pCtrl, strVal))
		{
			nPos = strVal->IndexOf(_strKey);
			if (nPos != -1)
				return;
			else
			{
				strVal += _strKey;
				m_CtrlEventDic[pCtrl] = strVal;
			}
		}
		else
		{
			m_CtrlEventDic->Add(pCtrl, _strKey);
		}
		switch (EventType)
		{
		case TangramClick:
			pCtrl->Click += gcnew EventHandler(this, &TangramProxy::OnClick);
			break;
		case TangramDoubleClick:
			pCtrl->DoubleClick += gcnew EventHandler(this, &TangramProxy::OnDoubleClick);
			break;
		case TangramEnter:
			pCtrl->Enter += gcnew EventHandler(this, &TangramProxy::OnEnter);
			break;
		case TangramLeave:
			pCtrl->Leave += gcnew EventHandler(this, &TangramProxy::OnLeave);
			break;
		case TangramEnabledChanged:
			pCtrl->EnabledChanged += gcnew EventHandler(this, &TangramProxy::OnEnabledChanged);
			break;
		case TangramLostFocus:
			pCtrl->LostFocus += gcnew EventHandler(this, &TangramProxy::OnLostFocus);
			break;
		case TangramGotFocus:
			pCtrl->GotFocus += gcnew EventHandler(this, &TangramProxy::OnGotFocus);
			break;
		case TangramKeyUp:
			pCtrl->KeyUp += gcnew KeyEventHandler(this, &TangramProxy::OnKeyUp);
			break;
		case TangramKeyDown:
			pCtrl->KeyDown += gcnew KeyEventHandler(this, &TangramProxy::OnKeyDown);
			break;
		case TangramKeyPress:
			pCtrl->KeyPress += gcnew KeyPressEventHandler(this, &TangramProxy::OnKeyPress);
			break;
		case TangramMouseClick:
			pCtrl->MouseClick += gcnew MouseEventHandler(this, &TangramProxy::OnMouseClick);
			break;// = 0x0000000a,
		case TangramMouseDoubleClick:
			pCtrl->MouseDoubleClick += gcnew MouseEventHandler(this, &TangramProxy::OnMouseDoubleClick);
			break;// = 0x0000000b,
		case TangramMouseDown:
			pCtrl->MouseDown += gcnew MouseEventHandler(this, &TangramProxy::OnMouseDown);
			break;// = 0x0000000c,
		case TangramMouseEnter:
			pCtrl->MouseEnter += gcnew EventHandler(this, &TangramProxy::OnMouseEnter);
			break;// = 0x0000000d,
		case TangramMouseHover:
			pCtrl->MouseHover += gcnew EventHandler(this, &TangramProxy::OnMouseHover);
			break;// = 0x0000000e,
		case TangramMouseLeave:
			pCtrl->MouseLeave += gcnew EventHandler(this, &TangramProxy::OnMouseLeave);
			break;// = 0x0000000f,
		case TangramMouseMove:
			pCtrl->MouseMove += gcnew MouseEventHandler(this, &TangramProxy::OnMouseMove);
			break;// = 0x00000010,
		case TangramMouseUp:
			pCtrl->MouseUp += gcnew MouseEventHandler(this, &TangramProxy::OnMouseUp);
			break;// = 0x00000011,
		case TangramMouseWheel:
			pCtrl->MouseWheel += gcnew MouseEventHandler(this, &TangramProxy::OnMouseWheel);
			break;// = 0x00000012,
		case TangramTextChanged:
			pCtrl->TextChanged += gcnew EventHandler(this, &TangramProxy::OnTextChanged);
			break;// = 0x00000013,
		case TangramVisibleChanged:
			pCtrl->VisibleChanged += gcnew EventHandler(this, &TangramProxy::OnVisibleChanged);
			break;// = 0x00000014,
		case TangramClientSizeChanged:
			pCtrl->SizeChanged += gcnew EventHandler(this, &TangramProxy::OnSizeChanged);
			break;// = 0x00000015,
		case TangramSizeChanged:
			pCtrl->ClientSizeChanged += gcnew EventHandler(this, &TangramProxy::OnClientSizeChanged);
			break;// = 0x00000016,
		case TangramParentChanged:
			pCtrl->ParentChanged += gcnew EventHandler(this, &TangramProxy::OnParentChanged);
			break;// = 0x00000017,
		case TangramResize:
			pCtrl->Resize += gcnew EventHandler(this, &TangramProxy::OnResize);
			break;// = 0x00000018
		}
	}

	void TangramProxy::DelegateEvent(Control^ pObj, IWndNode* pNode)
	{
		m_pCurrentNode = pNode;
		if (pNode)
		{
			IWndPage* pPage = nullptr;
			pNode->get_WndPage(&pPage);
			if (pPage)
			{
				VARIANT_BOOL b = false;
				pPage->get_EnableSinkCLRCtrlCreatedEvent(&b);
				if (b == TRUE)
				{
					pObj->HandleCreated += gcnew EventHandler(this,&TangramProxy::Obj_HandleCreated);
				}
			}
		}
	}

	void TangramProxy::Obj_HandleCreated(Object^ sender, EventArgs^ e)
	{
		Control^ pCtrl = (Control^)sender;
		HWND hWnd = (HWND)pCtrl->Handle.ToInt64();

		if (this->m_pCurrentNode)
		{
			long h = 0;
			m_pCurrentNode->get_Handle(&h);
			IWndFrame* pFrame = nullptr;
			m_pCurrentNode->get_Frame(&pFrame);
			pFrame->get_HWND(&h);
			CtrlInfo m_CtrlInfo;
			IntPtr nDisp;
			try
			{
				m_CtrlInfo.m_pCtrlDisp = (IDispatch*)Marshal::GetIDispatchForObject(pCtrl).ToPointer();
				//nDisp = (IntPtr)Marshal::GetIDispatchForObject(pCtrl).ToInt64();
			}
			catch (InvalidCastException^ e)
			{
				String^ s = e->Message;
			}
			finally
			{
				if (nDisp.ToInt64())
				{
					//m_CtrlInfo.m_pCtrlDisp = (IDispatch*)nDisp.ToPointer();
					m_CtrlInfo.m_hWnd = (HWND)pCtrl->Handle.ToInt64();
					m_CtrlInfo.m_strName = STRING2BSTR(pCtrl->Name);
					IWndPage* pPage = nullptr;
					m_pCurrentNode->get_WndPage(&pPage);
					m_CtrlInfo.m_pPage = pPage;
					m_CtrlInfo.m_pNode = m_pCurrentNode;
					::SendMessage((HWND)h, WM_TANGRAMMSG, (WPARAM)&m_CtrlInfo, 2048);;
				}
			}
		}

		Control::ControlCollection^ pControls = pCtrl->Controls;

		for (int i = 0; i < pControls->Count; i++)
		{
			Obj_HandleCreated(pControls[i], e);
		}
	}

	//void TangramProxy::Obj_HandleCreated(Object^ sender, EventArgs^ e)
	//{
	//	Control^ pCtrl = (Control^)sender;
	//	HWND hWnd = (HWND)pCtrl->Handle.ToInt64();
	//
	//	if (this->m_pCurrentNode)
	//	{
	//		LONGLONG h = 0;
	//		m_pCurrentNode->get_Handle(&h);
	//		CtrlInfo m_CtrlInfo;
	//		try
	//		{
	//			m_CtrlInfo.m_pCtrlDisp = (IDispatch*)Marshal::GetIDispatchForObject(pCtrl).ToPointer();
	//			m_CtrlInfo.m_hWnd = (HWND)pCtrl->Handle.ToInt64();
	//			m_CtrlInfo.m_strName = STRING2BSTR(pCtrl->Name);
	//			::SendMessage((HWND)h, WM_TANGRAMMSG, (WPARAM)&m_CtrlInfo, 2048);;
	//		}
	//		catch (InvalidCastException^ e)
	//		{
	//
	//		}
	//	}
	//
	//	Control::ControlCollection^ pControls = pCtrl->Controls;
	//
	//	for (int i = 0; i < pControls->Count; i++)
	//	{
	//		Obj_HandleCreated(pControls[i],e);
	//	}
	//}

	void TangramProxy::OnEvent(Object ^sender, EventArgs ^e, WindowEventType nType)
	{

	}

	void TangramProxy::OnClick(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramClick);
	}

	void TangramProxy::OnDoubleClick(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramDoubleClick);
	}

	void TangramProxy::OnEnter(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramEnter);
	}

	void TangramProxy::OnLeave(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramLeave);
	}

	void TangramProxy::OnEnabledChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramEnabledChanged);
	}

	void TangramProxy::OnLostFocus(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramLostFocus);
	}

	void TangramProxy::OnGotFocus(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramGotFocus);
	}

	void TangramProxy::OnKeyUp(Object ^sender, KeyEventArgs ^e)
	{
		OnEvent(sender, e, TangramKeyUp);
	}

	void TangramProxy::OnKeyDown(Object ^sender, KeyEventArgs ^e)
	{
		OnEvent(sender, e, TangramKeyDown);
	}

	void TangramProxy::OnKeyPress(Object ^sender, KeyPressEventArgs ^e)
	{
		OnEvent(sender, e, TangramKeyPress);
	}

	void TangramProxy::OnMouseClick(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseClick);
	}


	void TangramProxy::OnMouseDoubleClick(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseDoubleClick);
	}


	void TangramProxy::OnMouseDown(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseDown);
	}


	void TangramProxy::OnMouseEnter(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseEnter);
	}


	void TangramProxy::OnMouseHover(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseHover);
	}


	void TangramProxy::OnMouseMove(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseMove);
	}


	void TangramProxy::OnMouseLeave(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseLeave);
	}


	void TangramProxy::OnMouseUp(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseUp);
	}


	void TangramProxy::OnMouseWheel(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseWheel);
	}


	void TangramProxy::OnTextChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramTextChanged);
	}


	void TangramProxy::OnVisibleChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramVisibleChanged);
	}


	void TangramProxy::OnSizeChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramSizeChanged);
	}


	void TangramProxy::OnClientSizeChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramClientSizeChanged);
	}


	void TangramProxy::OnParentChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramParentChanged);
	}


	void TangramProxy::OnResize(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramResize);
	}

	void TangramProxy::OnHandleDestroyed(Object ^sender, EventArgs ^e)
	{
		Control^ pCtrl = ((Control^)sender);
		HWND hWnd = (HWND)pCtrl->Handle.ToInt64();
		ObjectEventInfo* pInfo = (ObjectEventInfo*)::GetWindowLongPtr(hWnd, GWLP_USERDATA);
		if (pInfo)
		{
			map<LONGLONG, ObjectEventInfo2*>::iterator it;
			for (it = pInfo->m_mapEventObj2.begin(); it != pInfo->m_mapEventObj2.end(); it++)
			{
				it->second->m_mapEventObj.clear();
				delete it->second;
			}
			pInfo->m_mapEventObj2.clear();
			delete pInfo;
		}
	}
}
