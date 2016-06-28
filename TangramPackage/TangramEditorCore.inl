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

// EditorCore.inl

/*****************************************************************************
** IVsWindowPane Implementation
*****************************************************************************/

template <class Traits_T>
STDMETHODIMP CAppXmlDoc<Traits_T>::CreatePaneWindow(
	_In_ HWND hWndParent,
	_In_ int x,
	_In_ int y,
	_In_ int cx,
	_In_ int cy,
	_Out_ HWND *phWnd)
{
	VSL_STDMETHODTRY{

	VSL_CHECKPOINTER(phWnd, E_INVALIDARG);

	RECT rect = { x, y, x+cx, y+cy };

	RichEditContainer::Create(hWndParent, rect);

	// Subclass the rich edit control window
	m_pWindowProc = reinterpret_cast<WNDPROC>(GetControl().SetWindowLongPtr(GWLP_WNDPROC, (LONG_PTR)CAppXmlDoc<Traits_T>::RichEditWindowProc));
	GetControl().SetWindowLongPtr(GWLP_USERDATA, (LONG_PTR)this);

	// Finally show the window
	GetControl().ShowWindow(SW_SHOWNORMAL);

	// Setup the selection container for the one selected item
	//GetItemsContainer()[0] = GetControl().GetITextDocument();

	// Now fire the selection change event
	//FireSelectionChange();

	// Register with the text manager, so find in files works correctly
	//RegisterToTextManager();

	*phWnd = GetHWND();

	}VSL_STDMETHODCATCH()

	return VSL_GET_STDMETHOD_HRESULT();
}

template <class Traits_T>
STDMETHODIMP CAppXmlDoc<Traits_T>::GetDefaultSize(_Out_ SIZE *psize)
{
	VSL_STDMETHODTRY{

	VSL_CHECKPOINTER(psize, E_INVALIDARG);

	psize->cx = 50;
	psize->cy = 50;

	}VSL_STDMETHODCATCH()

	return VSL_GET_STDMETHOD_HRESULT();
}

template <class Traits_T>
STDMETHODIMP CAppXmlDoc<Traits_T>::ClosePane()
{
	if(m_bClosed)
	{
		// recursion guard
		return E_UNEXPECTED;
	}

	VSL_STDMETHODTRY{

	// Put back the original window procedure
	GetControl().SetWindowLongPtr(GWLP_WNDPROC, (LONG_PTR)m_pWindowProc);

	// Destroy the parent window, which will in turn destroy the rich edit window
	Destroy();

	// Notify the persistence base class that the document is closing
	OnDocumentClose();

	m_bClosed = true;

	}VSL_STDMETHODCATCH()

	return VSL_GET_STDMETHOD_HRESULT();
}
