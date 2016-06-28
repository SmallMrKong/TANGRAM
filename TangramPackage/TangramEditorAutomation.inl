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

// EditorAutomation.inl

#pragma once

/*
ITangramEditor is the automation interface for CAppXmlDoc.  The implementation
of the methods is just a thin wrapper over the site cache and the rich edit control's object
model.
*/

// Window Proc for the rich edit control.  Necessary to implement macro recording.
template <class Traits_T>
LRESULT CALLBACK CAppXmlDoc<Traits_T>::RichEditWindowProc(
	_In_ HWND hWnd, 
	UINT msg, 
	_In_ WPARAM wParam, 
	_In_ LPARAM lParam)
{
	VSL_STDMETHODTRY{

	// Get the document pointer, which was set by calling SetWindowLong in CreatePaneWindow
	Control::Window window(hWnd);
	CAppXmlDoc* pDocument = reinterpret_cast<CAppXmlDoc*>(window.GetWindowLongPtr(GWLP_USERDATA));

	if(NULL != pDocument)
	{
		pDocument->EnsureNotClosed(); // paranoid, this shouldn't happen since this proc is removed in ClosePane
		switch(msg)
		{
		case WM_SETCURSOR:
			// Set the cursor to the "No" symbol, to indicate mouse actions will not be recorded, 
			// and suppress the message
			{
			static Traits_T::Cursor noCursor(IDC_NO);
			noCursor.Activate();
			}
			return 0;

		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
		case WM_MBUTTONDOWN:
		case WM_LBUTTONDBLCLK:
			// The cursor has been set to the "No" symbol, just set the focus, and suppress the message
			return 0;

		default:
			// Need to process the message normally, as well so that user action actually occurs in the 
			// control after recording the user action
			return pDocument->GetControl().CallWindowProc(pDocument->m_pWindowProc, msg, wParam, lParam);
		}
	}

	}VSL_STDMETHODCATCH()

	VSL_ASSERT(SUCCEEDED(VSL_GET_STDMETHOD_HRESULT()));

	return 0;
}

template <class Traits_T>
IDispatch* CAppXmlDoc<Traits_T>::GetNamedAutomationObject(_In_z_ BSTR bstrName)
{
	const wchar_t szDocumentName[] = L"Document";
	if(bstrName != NULL && bstrName[0] != L'\0')
	{
		// NULL or empty string just means the default object, but if a specific string
		// is specified, then make sure it's the correct one, but don't enforce case
		CHK(0 == ::_wcsicmp(bstrName, szDocumentName), E_INVALIDARG);
	}
	IDispatch* pIDispatch = static_cast<ITangramEditor*>(this);
	// Required to AddRef this once before returning it
	pIDispatch->AddRef();
	return pIDispatch;
}

