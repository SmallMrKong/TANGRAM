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

// EditorPersistence.inl
#include "def.h"
// Called by the IPersistFileFormat::Load implementation on DocumentPersistanceBase
template <class Traits_T>
HRESULT CAppXmlDoc<Traits_T>::ReadData(
	File& rFile, 
	BOOL bInsert,
	DWORD& rdwFormatIndex) throw()
{
	VSL_STDMETHODTRY{

	// Only the default format is supported by this editor
	rdwFormatIndex = DEF_FORMAT_INDEX;

	{

	Control::SuspendDrawAndNotifications suspend(GetControl());

	// Figure out format of file being read by examining the start of
	// the file for the RTF signature.
	const char szRTFSignature[] = "{\\rtf";
	char szHeader[_countof(szRTFSignature)];
	const DWORD dwBytesToRead = _countof(szRTFSignature) - 1; // -1 as the last spot is for the NULL terminator
	DWORD dwBytesRead;
	rFile.Read(szHeader, dwBytesToRead, dwBytesRead);

	// NULL terminate so this is a proper string
	szHeader[_countof(szHeader)-1] = '\0';

	// If the signature isn't RTF, then assume text
	DWORD dwFormat = SF_TEXT;
	if(dwBytesToRead == dwBytesRead && 0 == ::_strnicmp(szRTFSignature, szHeader, dwBytesToRead))
	{
		dwFormat = SF_RTF;
	}

	// Move back to the beginning of the file
	rFile.Seek(0L, FILE_BEGIN);

	// Now tell the control to load the file
	EDITSTREAM editStream =
	{
		reinterpret_cast<DWORD_PTR>(&rFile),
		S_OK,
		&EditStreamCallback<true>
	};

	// This message will result in EditStreamInCallback being called
	GetControl().SendMessage(
		EM_STREAMIN,
		(bInsert ? dwFormat | SFF_SELECTION : dwFormat),
		&editStream);
	if (theApp.m_pTangram)
	{
		CComPtr<ITextRange>	spITextRange;
		GetTextRange(spITextRange);
		CComBSTR bstr(L"");
		spITextRange->GetText(&bstr);

		CString strXml = OLE2T(bstr);
		if (theApp.m_pOutputWindowPane)
		{
			theApp.m_pOutputWindowPane->OutputString(bstr);
		}
		m_strData = strXml;
		theApp.m_pTangram->put_AppKeyValue(CComBSTR(L"TangramDesignerXml"), CComVariant(strXml));
		CTangramXmlParse m_Parse;
		if (m_Parse.LoadXml(theApp.m_strDesignerXml))
		{
			CString strUrl = m_Parse.attr(_T("URL"), _T(""));
			if (strUrl != _T(""))
			{
				HWND h = m_hWnd;
				theApp.m_pTangram->CreateWndPage((long)::GetParent(h), &m_pPage);
				if (m_pPage)
				{
					m_pPage->put_URL(strUrl.AllocSysString());
				}
			}
			else
			{
				CTangramXmlParse* pParse = m_Parse.GetChild(_T("window"));
				if (pParse)
				{
					HWND h = m_hWnd;
					theApp.m_pTangram->CreateWndPage((long)::GetParent(h), &m_pPage);
					if (m_pPage)
					{
						CComPtr<IWndFrame> m_pFrame;
						m_pPage->CreateFrame(CComVariant(0), CComVariant((long)h), CComBSTR(L""), &m_pFrame);
						m_pFrame->Extend(CComBSTR(L""), CComBSTR(theApp.m_strDesignerXml), &m_pNode);
						long h = 0;
						m_pNode->get_Handle(&h);
						m_pNode->Extend(CComBSTR(L""), strXml.AllocSysString(), &m_pDesignerNode);
						if (m_pDesignerNode)
						{
							IWndFrame* pFrame = NULL;
							m_pDesignerNode->get_Frame(&pFrame);
							pFrame->put_DesignerState(true);
						}
					}
				}
			}
		}
	}
	VSL_SET_STDMETHOD_HRESULT(*(reinterpret_cast<HRESULT*>(&editStream.dwError)));

	} // Suspend needs to be destroyed here

	if(SUCCEEDED(VSL_GET_STDMETHOD_HRESULT()))
	{
		// Redraw so that the new context is reflected on screen
		GetControl().InvalidateRect(NULL, true);
		GetControl().UpdateWindow();
		// Update the status bar, since the content is being loaded for the first time
		//SetInfo();
	}

	}VSL_STDMETHODCATCH()

	return VSL_GET_STDMETHOD_HRESULT();
}            

// Called indirectly by IPersistFileFormat::Save and IVsFileBackup::BackupFile implementations 
// on DocumentPersistanceBase
template <class Traits_T>
void CAppXmlDoc<Traits_T>::WriteData(File& rFile, DWORD /*dwFormatIndex*/)
{
	theApp.m_pTangram->UpdateWndNode(m_pDesignerNode);
	CComBSTR bstrXml(L"");
	m_pDesignerNode->get_DocXml(&bstrXml);
	CComPtr<ITextRange>	spITextRange;
	GetTextRange(spITextRange);
	if (spITextRange)
		spITextRange->SetText(bstrXml);

	EDITSTREAM editStream =
	{
		reinterpret_cast<DWORD_PTR>(&rFile),
		S_OK,
		&EditStreamCallback<false>
	};

	// This message will result in EditStreamCallback being called
	GetControl().SendMessage(EM_STREAMOUT, SF_TEXT, &editStream);
	//GetControl().SendMessage(EM_STREAMOUT, SF_RTF, &editStream);

	CHKHR(*(reinterpret_cast<HRESULT*>(&editStream.dwError)));
}

// Called by the Rich Edit control during the processing of EM_STREAMOUT and EM_STREAMIN
template <class Traits_T> // class template
template <bool bRead_T> // method template
DWORD CALLBACK CAppXmlDoc<Traits_T>::EditStreamCallback(
	_In_ DWORD_PTR dwpFile, 
	_Inout_bytecap_(iBufferByteSize) LPBYTE pBuffer, 
	LONG iBufferByteSize, 
	_Out_ LONG* piBytesWritenOrRead)
{
	DWORD dwBytesWritten = 0;

	DWORD dwBufferByteSize = iBufferByteSize;

	VSL_STDMETHODTRY{

	if(	NULL != pBuffer && 
		0 != dwBufferByteSize && 
		LONG_MAX >= dwBufferByteSize && 
		NULL != piBytesWritenOrRead)
	{
		File* pFile = reinterpret_cast<File*>(dwpFile);

		if(NULL != pFile)
		{
			if(bRead_T)
			{
				pFile->Read(pBuffer, dwBufferByteSize, dwBytesWritten);
			}
			else
			{
				pFile->Write(pBuffer, dwBufferByteSize, &dwBytesWritten);
			}
			VSL_ASSERT(dwBytesWritten <= iBufferByteSize);
		}

		if(dwBytesWritten <= dwBufferByteSize)
		{
			*piBytesWritenOrRead = dwBytesWritten;
		}
	}

	}VSL_STDMETHODCATCH()

	HRESULT hr = (dwBytesWritten >= 0 && (bRead_T ? true : dwBytesWritten == dwBufferByteSize)) ? 
		VSL_GET_STDMETHOD_HRESULT() : 
		(FAILED(VSL_GET_STDMETHOD_HRESULT()) ? VSL_GET_STDMETHOD_HRESULT() : E_FAIL);

	// If the return value is non-Zero (i.e. not S_OK), the return value will
	// be put into the dwError member of the EDITSTREAM instance passed with
	// the EM_STREAMOUT message in WriteData
	return *(reinterpret_cast<DWORD*>(&hr));
}

// Called by VSL::DocumentPersistanceBase::InitNew and VSL::DocumentPersistanceBase::Save
template <class Traits_T>
bool CAppXmlDoc<Traits_T>::IsValidFormat(DWORD dwFormatIndex)
{
	// Only one format, the default, is supported
	return DEF_FORMAT_INDEX == dwFormatIndex;
}

// Called by VSL::DocumentPersistanceBase::FilesChanged
template <class Traits_T>
void CAppXmlDoc<Traits_T>::OnFileChangedSetTimer()
{
	// 500 milliseconds is an arbitrary time to delay
	// See CAppXmlDoc::OnTimer and 
	// VSL::DocumentPersistanceBase::FilesChanged for details
	VSL_CHECKBOOL_GLE(0 != SetTimer(WFILECHANGEDTIMERID, 500, NULL));
}

// Called by VSL::DocumentPersistanceBase::GetClassID, which is also called by 
// VSL::DocumentPersistanceBase::GetGuidEditorType)
template <class Traits_T>
const GUID& CAppXmlDoc<Traits_T>::GetEditorTypeGuid() const
{
	// The GUID for the factory is the one to return from IPersist::GetClassID and 
	// IVsPersistDocData::GetGuidEditorType
	return CLSID_TangramPackageEditorFactory;
}

// Called by VSL::DocumentPersistanceBase::GetFormatList
template <class Traits_T>
void CAppXmlDoc<Traits_T>::GetFormatListString(ATL::CStringW& rstrFormatList)
{
	// Load the file format list string from the resource DLL
	VSL_CHECKBOOL_GLE(rstrFormatList.LoadString(IDS_FORMATSTR));
}

// Called indirectly by VSL::DocumentPersistanceBase::Load and VSL::DocumentPersistanceBase::Save
template <class Traits_T>
void CAppXmlDoc<Traits_T>::PostSetDirty()
{
	// Notify the Rich Edit control of the current dirty state
	GetControl().SetModified(IsFileDirty());
}

// Called indirectly by VSL::DocumentPersistanceBase::FilesChanged, 
// VSL::DocumentPersistanceBase::IgnoreFileChanges, VSL::DocumentPersistanceBase::Load, and 
// VSL::DocumentPersistanceBase::Save
template <class Traits_T>
void CAppXmlDoc<Traits_T>::PostSetReadOnly()
{
	// Notify the Rich Edit control of the current read-only state.
	GetControl().SetReadOnly(IsFileReadOnly());

	// Update editor caption with " [Read Only]" or "" as necessary

	CComPtr<IVsWindowFrame> spIVsWindowFrame;
	CHKHR(GetVsSiteCache().QueryService(SID_SVsWindowFrame, &spIVsWindowFrame));

	CStringW strEditorCaption;

	if(IsFileReadOnly())
	{
		VSL_CHECKBOOL_GLE(strEditorCaption.LoadString(IDS_READONLY));
	}

	CComVariant varEditorCaption = strEditorCaption;

	CHKHR(spIVsWindowFrame->SetProperty(VSFPROPID_EditorCaption, varEditorCaption));
}

template <class Traits_T>
void CAppXmlDoc<Traits_T>::UpdateVSCommandUI()
{
	CComPtr<IVsUIShell> spIVsUIShell = GetVsSiteCache().GetCachedService<IVsUIShell, SID_SVsUIShell>();
	if (!!spIVsUIShell)
	{
		// Tell VS to update the menus and toolbars
		CHKHR(spIVsUIShell->UpdateCommandUI(false));
	}
}
