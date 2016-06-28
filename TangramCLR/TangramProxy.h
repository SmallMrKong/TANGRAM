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

using namespace System;
using namespace System::ComponentModel;
using namespace System::Collections;
using namespace System::Collections::Generic;
using namespace System::Diagnostics;
using namespace System::Windows::Forms;

using namespace System::IO;

namespace TangramCLR {
	/// <summary>
	/// Summary for TangramProxy
	/// </summary>
	private ref class TangramProxy
	{
	public:
		TangramProxy(void);

		//HWND m_hParentWnd;
		IWndNode*					m_pCurrentNode;
		Dictionary<Control^, String^>^	m_CtrlEventDic;
		void DelegateEvent(Control^ pObj,IWndNode* pNode);
		void AttachEvent(Control^ pCtrl, WindowEventType EventType);
		void AttachHandleDestroyedEvent(Control^ pCtrl);
		void OnEvent(System::Object ^sender, System::EventArgs ^e, WindowEventType nType);
	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~TangramProxy()
		{
		}

	private:
		void Obj_HandleCreated(Object^ sender, EventArgs^ e);
	
		void OnKeyPress(System::Object ^sender, System::Windows::Forms::KeyPressEventArgs ^e);
		void OnClick(System::Object ^sender, System::EventArgs ^e);
		void OnDoubleClick(System::Object ^sender, System::EventArgs ^e);
		void OnEnter(System::Object ^sender, System::EventArgs ^e);
		void OnLeave(System::Object ^sender, System::EventArgs ^e);
		void OnEnabledChanged(System::Object ^sender, System::EventArgs ^e);
		void OnLostFocus(System::Object ^sender, System::EventArgs ^e);
		void OnGotFocus(System::Object ^sender, System::EventArgs ^e);
		void OnKeyUp(System::Object ^sender, System::Windows::Forms::KeyEventArgs ^e);
		void OnKeyDown(System::Object ^sender, System::Windows::Forms::KeyEventArgs ^e);
		void OnMouseClick(System::Object ^sender, System::Windows::Forms::MouseEventArgs ^e);
		void OnMouseDoubleClick(System::Object ^sender, System::Windows::Forms::MouseEventArgs ^e);
		void OnMouseDown(System::Object ^sender, System::Windows::Forms::MouseEventArgs ^e);
		void OnMouseEnter(System::Object ^sender, System::EventArgs ^e);
		void OnMouseHover(System::Object ^sender, System::EventArgs ^e);
		void OnMouseMove(System::Object ^sender, System::Windows::Forms::MouseEventArgs ^e);
		void OnMouseLeave(System::Object ^sender, System::EventArgs ^e);
		void OnMouseUp(System::Object ^sender, System::Windows::Forms::MouseEventArgs ^e);
		void OnMouseWheel(System::Object ^sender, System::Windows::Forms::MouseEventArgs ^e);
		void OnTextChanged(System::Object ^sender, System::EventArgs ^e);
		void OnVisibleChanged(System::Object ^sender, System::EventArgs ^e);
		void OnSizeChanged(System::Object ^sender, System::EventArgs ^e);
		void OnClientSizeChanged(System::Object ^sender, System::EventArgs ^e);
		void OnParentChanged(System::Object ^sender, System::EventArgs ^e);
		void OnResize(System::Object ^sender, System::EventArgs ^e);
		void OnHandleDestroyed(System::Object ^sender, System::EventArgs ^e);
	};
}
