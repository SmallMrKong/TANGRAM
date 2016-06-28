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

// TangramEditorDocument.h

#pragma once

/***************************************************************************

EditorDocument provides the implementation of a single view editor (as opposed to a 
multi-view editor that can display the same file in multiple modes like the HTML editor).

TangramPackage.pkgdef contains:

[$RootKey$\KeyBindingTables\{7f8286c6-672d-484a-8c4b-1f257d7353b1}]
@="#1"
"AllowNavKeyBinding"=dword:00000000
"Package"="{19631222-1992-0612-1965-060119821989}"

which is required for some, but not all, of the key bindings, which are located at the bottom of 
TangramPackageUI.vsct, to work correctly so that the appropriate command handler below will be
called.

***************************************************************************/
// This is provided as a template parameter to facilitate unit testing
// One traits class is provided rather then multiple template arguments
// so that when an new type is added, it isn't necessary to update all
// instances of the template declaration

#include "def.h"
#include "TangramPackage.h"

class CAppXmlDocDefault
{
public:
	// The rich edit control needs to be contained in another window, as the list view will send 
	// it's notifications to it's parent window; however, if the VS frame window is subclassed, VS
	// may stomp on the window proc we have set, so we have to provide our own parent window
	// for the control.
	typedef Win32ControlContainer<RichEditWin32Control<> > RichEditContainer;
	typedef VSL::Cursor Cursor;
	typedef VSL::Keyboard Keyboard;
	typedef VSL::File File;
};

template <class Traits_T = CAppXmlDocDefault>
class CAppXmlDoc : 
	public Traits_T::RichEditContainer,
	public CComObjectRootEx<CComSingleThreadModel>,
	public IVsWindowPaneImpl<CAppXmlDoc<Traits_T> >,
	public IExtensibleObjectImpl<CAppXmlDoc<Traits_T> >,
	public IDispatchImpl<ITangramEditor, &__uuidof(ITangramEditor)>,
	public DocumentPersistanceBase<CAppXmlDoc<Traits_T>, typename Traits_T::File>
{

// COM objects typically should not be cloned, and this prevents cloning by declaring the 
// copy constructor and assignment operator private (NOTE:  this macro includes the declaration of
// a private section, so everything following this macro and preceding a public or protected 
// section will be private).
VSL_DECLARE_NOT_COPYABLE(CAppXmlDoc)

public:
	IWndPage* m_pPage;
	IWndNode* m_pNode;
	IWndNode* m_pDesignerNode;
	CString m_strKey;
	CString m_strData;
	typedef typename Traits_T::File File;

// Provides a portion of the implementation of IUnknown, in particular the list of interfaces
// the CAppXmlDoc object will support via QueryInterface
BEGIN_COM_MAP(CAppXmlDoc)
	COM_INTERFACE_ENTRY(IPersist)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IVsWindowPane)    
	COM_INTERFACE_ENTRY(IVsPersistDocData)
	COM_INTERFACE_ENTRY(IExtensibleObject)
	COM_INTERFACE_ENTRY(ITangramEditor)
	COM_INTERFACE_ENTRY(IPersistFileFormat)
	COM_INTERFACE_ENTRY(IVsFileChangeEvents)
	COM_INTERFACE_ENTRY(IVsDocDataFileChangeControl)
END_COM_MAP()

VSL_BEGIN_MSG_MAP(CAppXmlDoc)
	// Whenever the content changes, need to check and see if the file can be edited,
	// if not then the changes will be rejected.
	COMMAND_HANDLER(RichEditContainer::iContainedControlID, EN_CHANGE, OnContentChange)
	// Whenever the selection changes, need to update the Visual Studio UI (i.e. status bar, menus, 
	// and toolbars).
	//NOTIFY_HANDLER(RichEditContainer::iContainedControlID, EN_SELCHANGE, OnSelectionChange)
	// On this event the context menu will be shown if needed, and some keyboard commands will be
	// dealt with.
	NOTIFY_HANDLER(RichEditContainer::iContainedControlID, EN_MSGFILTER, OnUserInteractionEvent)
	MESSAGE_HANDLER(WM_TIMER, OnTimer)
	MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
	MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
	MESSAGE_HANDLER(WM_TANGRAMMSG, OnTangramMsg)
	// Let the rich edit container process all other messages
	CHAIN_MSG_MAP(RichEditContainer)
VSL_END_MSG_MAP()

// This is necessary to implement the connection point for IVsTextViewEvents
// See comment above SingleViewFindInFilesOutputWindowIntegrationImpl
BEGIN_CONNECTION_POINT_MAP(CAppXmlDoc)
	CONNECTION_POINT_ENTRY(IID_IVsTextViewEvents)
END_CONNECTION_POINT_MAP()

// IVsWindowPane methods overridden or not provided by IVsWindowPaneImpl

	STDMETHOD(CreatePaneWindow)(
		/*[in] */ _In_ HWND hwndParent,
		/*[in] */ _In_ int x,
		/*[in] */ _In_ int y,
		/*[in] */ _In_ int cx,
		/*[in] */ _In_ int cy,
		/*[out]*/ _Out_ HWND *hwnd);
	STDMETHOD(GetDefaultSize)(
		/*[out]*/ _Out_ SIZE *psize);
	STDMETHOD(ClosePane)();

	IDispatch* GetNamedAutomationObject(_In_z_ BSTR bstrName);
	bool IsValidFormat(DWORD dwFormatIndex);
	void OnFileChangedSetTimer();
	const GUID& GetEditorTypeGuid() const;
	void GetFormatListString(ATL::CStringW& rstrFormatList);
	void PostSetDirty();
	void PostSetReadOnly();
	HRESULT ReadData(File& rFile, BOOL bInsert, DWORD& rdwFormatIndex) throw();
	void WriteData(File& rFile, DWORD /*dwFormatIndex*/);

protected:

	typedef typename Traits_T::RichEditContainer RichEditContainer;
	typedef typename Traits_T::RichEditContainer::Control Control;
	typedef typename Traits_T::Keyboard Keyboard;

	typedef typename IVsWindowPaneImpl<CAppXmlDoc<Traits_T> >::VsSiteCache VsSiteCache;

	typedef CAppXmlDoc<Traits_T> This;

	CAppXmlDoc():
		m_pWindowProc(NULL),
		m_bClosed(false)
	{
		m_strKey		= _T("");
		m_strData		= _T("");
		m_pPage			= nullptr;
		m_pNode			= nullptr;
		m_pDesignerNode = nullptr;
	}

	~CAppXmlDoc()
	{
		VSL_ASSERT(m_bClosed);
		if(!m_bClosed)
		{
			ClosePane();
		}
	}

	// Called by the Rich Edit control during the processing of EM_STREAMOUT and EM_STREAMIN
	template <bool bRead_T>
	static DWORD CALLBACK EditStreamCallback(
		_In_ DWORD_PTR dwpFile, 
		_Inout_bytecap_(iBufferByteSize) LPBYTE pBuffer, 
		LONG iBufferByteSize, 
		_Out_ LONG* piBytesWritten);

	// Window Proc for the rich edit control.  Necessary to implement macro recording.
	static LRESULT CALLBACK RichEditWindowProc(
		_In_ HWND hWnd, 
		UINT msg, 
		_In_ WPARAM wParam, 
		_In_ LPARAM lParam);

#pragma warning(push)
#pragma warning(disable : 4480) // // warning C4480: nonstandard extension used: specifying underlying type for enum
	enum TimerID : WPARAM
	{
		// ID of timer message sent from OnFileChangedSetTimer
		WFILECHANGEDTIMERID = 1,
		// ID of timer message sent from OnSetFocus
		WDELAYSTATUSBARUPDATETIMERID = 2,
	};
#pragma warning(pop)

private:

// Windows message handlers
	LRESULT OnContentChange(WORD /*iCommand*/, WORD /*iId*/, _In_ HWND hWindow, BOOL& /*bHandled*/);
	LRESULT OnUserInteractionEvent(int /*wParam*/, _In_ LPNMHDR pHeader, BOOL& /*bHandled*/);
	LRESULT OnTimer(UINT /*uMsg*/, _In_ WPARAM wParam, _In_ LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocus(UINT /*uMsg*/, _In_ WPARAM wParam, _In_ LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnDestroy(UINT /*uMsg*/, _In_ WPARAM wParam, _In_ LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnTangramMsg(UINT, WPARAM, LPARAM, BOOL&);

	// Rich Edit object model helper methods

	void GetTextSelection(CComPtr<ITextSelection>& rspITextSelection)
	{
		EnsureNotClosed();
		ITextDocument* pdoc = GetControl().GetITextDocument();
		CHKPTR(pdoc, E_FAIL);
		CHKHR(pdoc->GetSelection(&rspITextSelection));
		CHKPTR(rspITextSelection.p, E_FAIL);
	}

	void GetTextRange(CComPtr<ITextRange>& rspITextRange)
	{
		EnsureNotClosed();
		ITextDocument* pITextDocument = NULL;
		pITextDocument = GetControl().GetITextDocument();
		if(pITextDocument)
		{ 
			CHKHR(pITextDocument->Range(0, tomForward, &rspITextRange));
			CHKPTR(rspITextRange.p, E_FAIL);
		}
	}

	//// Helper function to checkout file if necessary.
	bool ShouldDiscardChange();

	// Calls IVsUIShell method to tell the environment to update the state of 
	// the command bars (menus and toolbars).
	void UpdateVSCommandUI();

	void EnsureNotClosed()
	{
		CHK(!m_bClosed, E_UNEXPECTED);
	}

	// Window procedure pointer for the window containing the rich edit control
	WNDPROC m_pWindowProc;

	// IVsWindowPane Data
	bool m_bClosed;
};
