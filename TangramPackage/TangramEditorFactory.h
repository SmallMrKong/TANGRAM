﻿/********************************************************************************
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

// TangramEditorFactory.h

#pragma once

#include "TangramEditorDocument.h"

/***************************************************************************
TangramPackage.pkgdef contains:


[$RootKey$\Editors\{19631222-1992-0612-1965-060119822016}]
@="Tangram"
"Package"="{19631222-1992-0612-1965-060119821989}"
"DisplayName"="Tangram"
"ExcludeDefTextEditor"=dword:00000001
"AcceptBinaryFiles"=dword:00000000

[$RootKey$\Editors\{19631222-1992-0612-1965-060119822016}\LogicalViews]
"{7651A703-06E5-11D1-8EBD-00A0C90F26EA}"=""

[$RootKey$\Editors\{19631222-1992-0612-1965-060119822016}\Extensions]
"tangram"=dword:00000032

which informs the shell that TangramPackagePackage is the package to create to have this editor 
registered with Visual Studio.  When TangramPackagePackage is sited by Visual Studio, 
TangramPackagePackage::PostSited creates an instance of CAppXmlDocEditorFactory and then queries
for SID_SVsRegisterEditors and then calls IVsRegisterEditors->RegisterEditor to register
the new instance of CAppXmlDocEditorFactory with Visual Studio.

Visual Studio then uses the CEditorFactory instance to create EditorDocument
instances, which is the actual editor instance.
***************************************************************************/

class CAppXmlDocEditorFactory :
	public CComObjectRootEx<CComSingleThreadModel>,
	public IVsEditorFactoryImpl<CAppXmlDocEditorFactory>
{

// Provides a portion of the implementation of IUnknown, in particular the list of interfaces
// the CEditorFactory object will support via QueryInterface
BEGIN_COM_MAP(CAppXmlDocEditorFactory)
    COM_INTERFACE_ENTRY(IVsEditorFactory)
END_COM_MAP()

// COM objects typically should not be cloned, and this prevents cloning by declaring the 
// copy constructor and assignment operator private (NOTE:  this macro includes the declaration of
// a private section, so everything following this macro and preceding a public or protected 
// section will be private).
VSL_DECLARE_NOT_COPYABLE(CAppXmlDocEditorFactory)

protected:
	CAppXmlDocEditorFactory()
	{
	}

	virtual ~CAppXmlDocEditorFactory()
	{
	}

public:

#pragma warning(push)
#pragma warning(disable : 4480) // // warning C4480: nonstandard extension used: specifying underlying type for enum
	enum PhysicalViewId : unsigned int
	{
		Unsupported,
		Primary
	};
#pragma warning(pop)

	PhysicalViewId GetPhysicalViewId(REFGUID rguidLogicalView)
	{
		if(LOGVIEWID_Primary == rguidLogicalView || LOGVIEWID_TextView == rguidLogicalView)
		{
			return Primary;
		}

		return Unsupported;
	}

	BSTR GetPhysicalViewBSTR(PhysicalViewId viewId)
	{
		if(Primary == viewId)
		{
			// Note that the value of LOGVIEWID_TextView in the registry is empty string.
			// The values under LogicalViews for an editor should match up to the values 
			// returned by this method, so that GetPhysicalViewId below can properly
			// map the strings back to a PhysicalViewId.
			return ::SysAllocString(L"");
		}

		return NULL;
	}

	PhysicalViewId GetPhysicalViewId(LPCOLESTR szPhysicalView)
	{
		if(szPhysicalView != NULL && szPhysicalView[0] == L'\0')
		{
			return Primary;
		}

		return Unsupported;
	}

	bool CanShareBuffer(PhysicalViewId /*physicalViewId*/)
	{
		return false;
	}

	void CreateSingleViewObject(
		PhysicalViewId physicalViewId, 
		CComPtr<IUnknown>& rspViewObject, 
		CComBSTR& rbstrEditorCaption, 
		const GUID*& rpguidCommandUI, 
		VSEDITORCREATEDOCWIN& /*rCreateDocumentWindowUI*/)
	{
		if(physicalViewId == Primary)
		{
			CComObject<CAppXmlDoc<> > *pDocument;
			VSL_CHECKHRESULT(CComObject<CAppXmlDoc<> >::CreateInstance(&pDocument));
			HRESULT hr = pDocument->QueryInterface(&rspViewObject);
			if(FAILED(hr))
			{
				// If QueryInterface failed, then there is something wrong with the object.
				// Delete it and throw an exception for the error.
				delete pDocument;
				VSL_CREATE_ERROR_HRESULT(hr);
			}

			// NOTE - if the file is read only [Read Only] will be appended to the caption
			rbstrEditorCaption.LoadStringW(IDS_EDITORCAPTION);

			rpguidCommandUI = &CLSID_TangramPackageEditorDocument;

			// pCreateDocumentWindowUI will be initialized to 0, and see we provide
			// no GUI for the user to cancel from so leave it 0

			return;
		}
		ERRHR(E_FAIL); // This should never happen
	}

private:

};
