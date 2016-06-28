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
// TangramEditorEvents.inl

/*
NOTE - pOleCmd is not ensured against NULL in the OnQuery* methods, as 
VSL::CommandHandlerBase::QueryStatus, which is the called of the OnQuery* method implemented 
here, already does so.  Additionally, VSL::IOleCommandTargetImpl::QueryStatus checks
pOleCmd as well before calling the query status handler.

The OnQuery* are all private, so the only way to access them is via calling QueryStatus.
*/

#include <VSLFont.h>
#include "def.h"
#include <map>
template <class Traits_T>
LRESULT CAppXmlDoc<Traits_T>::OnContentChange(WORD /*iCommand*/, WORD /*iId*/, _In_ HWND hWindow, BOOL& /*bHandled*/)
{
	CHK(hWindow == GetControl().GetHWND(), E_FAIL);

	// If not dirty then it is necessary to process this message in order to determine if
	// the file should be marked as dirty or the change undone
	if(!IsFileDirty())
	{
		// Check with the source control provider to see if the file can be changed.
		if(!ShouldDiscardChange())
		{
			SetFileDirty(true);
			// REVIEW - 3/14/2006 - is this here to ensure the "*" get's appended to the document name properly?
			UpdateVSCommandUI();    
		}
		else
		{
			// The source control provider indicate that this file can't be edited, so undo the changes.
			GetControl().Undo();
		}
	}
	return 0; // return value is ignored
}

template <class Traits_T>
LRESULT CAppXmlDoc<Traits_T>::OnUserInteractionEvent(int /*wParam*/, _In_ LPNMHDR pHeader, BOOL& /*bHandled*/)
{
	CHKPTR(pHeader, E_FAIL);
	MSGFILTER* pMsgFilter = reinterpret_cast<MSGFILTER*>(pHeader);

	// Non-zero indicates the RTF control should not handle the message
	// Use this as the default and let the default cases set it to 0
	// for events not handled here
	LRESULT iRet = 1;

	switch(pMsgFilter->msg)
	{
	case WM_RBUTTONUP:
		{
			// Convert to screen coordinates
			POINT pt;
			POINTSTOPOINT(pt, MAKEPOINTS(pMsgFilter->lParam));
			GetControl().ClientToScreen(&pt);
			CHK(pt.x <= SHRT_MAX && pt.x > 0, E_FAIL);
			CHK(pt.y <= SHRT_MAX && pt.y > 0, E_FAIL);

			POINTS ptsTemp = {static_cast<short>(pt.x) , static_cast<short>(pt.y)};

			// Now show the context menu
			CComPtr<IOleComponentUIManager> spIOleComponentUIManager;
			CHKHR(GetVsSiteCache().QueryService(SID_SOleComponentUIManager, &spIOleComponentUIManager));
			CHKHR(spIOleComponentUIManager->ShowContextMenu(OLEROLE_TOPLEVELCOMPONENT,
			   CLSID_TangramPackageCmdSet,
			   IDMX_RTF,
			   ptsTemp,
			   (IOleCommandTarget *) this));
		}
		break;               

	case WM_CHAR:
		{

		// CTRL+A/"Select All" is handled separately from the others, since this command does
		// does not modify the content
		if(1 == pMsgFilter->wParam)
		{
			iRet = 0;
			break;
		}

		// If the document is dirty, it isn't necessary to determine if the file
		// can be edited currently, as it has to editable in order to be dirty
		bool bDiscardInput = false;
		if(!IsFileDirty())
		{
			// Determine if editing is currently possible
			bDiscardInput = ShouldDiscardChange();
		}

		// Discard the input if the user has decided not to check the file out
		// or if the file has been reloaded
		if(bDiscardInput || IsFileReloaded())
		{
			SetFileReloaded(false);
			break;
		}

		iRet = 0;

		break;
		}

	default:
		// Return value of 0 indicates that the control should process the event.
		// Since this message isn't handled here, let the control handle it.
		iRet = 0;
		break;
	}

	
	return iRet;
}

template <class Traits_T>
LRESULT CAppXmlDoc<Traits_T>::OnTimer(UINT /*uMsg*/, _In_ WPARAM wParam, _In_ LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    if(WFILECHANGEDTIMERID == wParam)
    {
        // This timer is set in DocumentPersistanceBase::FilesChanged which is a notification
        // from the SVsFileChangeEx service that the file has changed outside of the 
		// environment. Prompt the user as to whether they want to reload the file or not.
        // See comments in DocumentPersistanceBase::FilesChanged for more details.

		Control::Window activeWindow(Control::Window::GetActiveWindow());
		if(!activeWindow.IsWindow())
		{
			// GetActiveWindow doesn't set the last error, so all we do here is indicate
			// that the message was processed (nothing else with process this message,
			// correctly either).
			return 0; // 0 indicates the message was processed
		}
        
        // If this process is the active window, shut off the timer and prompt to reload file.
		if(::GetCurrentProcessId() == GetWindowProcessID())
        {
            KillTimer(WFILECHANGEDTIMERID);

			OleComponentUIManagerUtilities<VsUtilityLocalSiteControl> util;
			util.SetSite(GetVsSiteCache());
			if(IDYES == util.ShowMessage(const_cast<wchar_t*>(static_cast<const wchar_t*>(GetFileName())), IDS_OUTSIDEEDITORFILECHANGE, OLEMSGBUTTON_YESNO, OLEMSGDEFBUTTON_FIRST, OLEMSGICON_QUERY))
            {
                ReloadDocData(0);
            }

            NotifyFileChangedTimerHandled();
        }
    }
	else if(WDELAYSTATUSBARUPDATETIMERID == wParam)
	{
		KillTimer(WDELAYSTATUSBARUPDATETIMERID);
		//SetInfo();
	}
	return 0; // 0 indicates the message was processed
}

template <class Traits_T>
LRESULT CAppXmlDoc<Traits_T>::OnTangramMsg(UINT uMsg, _In_ WPARAM wParam, _In_ LPARAM lParam, BOOL& bHandled)
{
	CString strKey = (LPCTSTR)wParam;
	if (strKey.CompareNoCase(_T("default")) == 0)
	{
		HWND h = m_hWnd;
		// Just hook this message, don't handle it.
		bHandled = false;

		CComPtr<ITextRange>	spITextRange;
		GetTextRange(spITextRange);
		if (spITextRange)
		{
			CComBSTR bstr(L"");
			spITextRange->GetText(&bstr);

			CString strXml = OLE2T(bstr);
			if (m_pNode)
			{
				IWndFrame* pFrame = NULL;
				m_pNode->get_Frame(&pFrame);
				pFrame->put_FrameData(CComBSTR(L"default"), CComVariant(strXml));
			}
		}
	}

	return DefWindowProc(uMsg, wParam, lParam); 
}

template <class Traits_T>
LRESULT CAppXmlDoc<Traits_T>::OnDestroy(UINT uMsg, _In_ WPARAM wParam, _In_ LPARAM lParam, BOOL& bHandled)
{
	theApp.m_pTangram->put_AppKeyValue(CComBSTR(L"TangramDesignerXml"), CComVariant(L""));
	return DefWindowProc(uMsg, wParam, lParam);
}

template <class Traits_T>
LRESULT CAppXmlDoc<Traits_T>::OnSetFocus(UINT /*uMsg*/, _In_ WPARAM /*wParam*/, _In_ LPARAM /*lParam*/, BOOL& bHandled)
{	
	//theApp.m_pTangramPackage->CreateTool(CLSID_guidPersistanceSlot);
	if (m_pNode)
		theApp.m_pTangram->put_CurrentDesignNode(m_pNode);
	HWND h = m_hWnd;
	CString strXml = _T("");
	// Just hook this message, don't handle it.
	bHandled = false;
	if (m_pDesignerNode)
	{
		CComBSTR bstrXml(L"");
		m_pDesignerNode->get_DocXml(&bstrXml);
		CComPtr<ITextRange>	spITextRange;
		GetTextRange(spITextRange);
		if(spITextRange)
			spITextRange->SetText(bstrXml);
		strXml = OLE2T(bstrXml);
	}

	if (strXml == _T(""))
	{
		CComPtr<ITextRange>	spITextRange;
		GetTextRange(spITextRange);
		if (spITextRange)
		{
			CComBSTR bstr(L"");
			spITextRange->GetText(&bstr);

			strXml = OLE2T(bstr);
		}
	}
	if (theApp.m_pOutputWindowPane)
	{
		theApp.m_pOutputWindowPane->Clear();
		theApp.m_pOutputWindowPane->OutputString(CComBSTR(strXml));
	}
	theApp.m_pTangram->put_AppKeyValue(CComBSTR(L"TangramDesignerXml"), CComVariant(strXml));

	// The results pane is very aggresive in updating the status bar
	// and will update it even after it has lost focus, so set a short
	// timer to update the staus bar again, in case this occurs
	VSL_CHECKBOOL_GLE(0 != SetTimer(WDELAYSTATUSBARUPDATETIMERID, 100, NULL));
	return 0;  // Ignored, since this is not the final handler of the message
}

// Only the Copy, Cut, Paste and Delete commands are explicitly recorded by the document.  The 
// command will be recorded via the macro default recording mechanism that attempts to record 
// any command not handled by this document

#define RETURN_IF_CANT_EDIT() \
	if(ShouldDiscardChange()) \
	{ \
		return; \
	}

template <class Traits_T>
bool CAppXmlDoc<Traits_T>::ShouldDiscardChange()
{
	// If the buffer is already Read/Write or edit enabled no need to check further
	// as the change need not be discarded
	if(IsFileEditableWhenReadOnly())
	{
		return false;
	}

    CComPtr<IVsQueryEditQuerySave2> spIVsQueryEditQuerySave2;
	CHKHR(GetVsSiteCache().QueryService(SID_SVsQueryEditQuerySave, &spIVsQueryEditQuerySave2));

	// Array of the file names to check, which in this case is just one.
	LPCOLESTR bstrFiles[] = {GetFileName()};
	VSQueryEditResult iEditResult = QER_EditNotOK;
	VSQueryEditResultFlags iEditResultFlags = 0;
	CHKHR(spIVsQueryEditQuerySave2->QueryEditFiles(
		QEF_AllowInMemoryEdits,
		1,
		bstrFiles,
		NULL,
		NULL,
		&iEditResult,
		&iEditResultFlags));

	if(QER_EditOK == iEditResult)
	{
		GetControl().SetReadOnly(false);

		// We cannot distinguish here between the case when a file has been checked out and the
		// case where the user has decided to allow in-memory edits for a read-only file so we
		// set "read edits enabled" bit.
		if(iEditResultFlags & QER_InMemoryEdit)
		{
			// QER_InMemoryEdit indicates an in memory edit for a read-only file
			SetFileEditableWhenReadOnly(true);
		}

		return false;
	}

	// In this case, the file is still read only, so the change needs to be discarded.
	return true;
}

