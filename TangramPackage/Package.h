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
// Package.h

#pragma once

#include "Guids.h"
#include <commctrl.h>
#include "resource.h"       // main symbols

#include "..\TangramPackageUI\Resource.h"
#include "..\TangramPackageUI\CommandIds.h"

#include "TangramToolWindow.h"
#include "TangramEditorFactory.h"


using namespace VSL;

/***************************************************************************
CAppDocPackage handles the necessary registration for this package.

See TangramEditorFactory.h for the details of the Editor key section in 
TangramPackage.pkgdef.

See the Package C++ reference sample for the details of the Package key section in
TangramPackage.pkgdef.

See the MenuAndCommands C++ reference sample for the details of the Menu key section in 
TangramPackage.pkgdef.

See TangramEditorDocument.h for the details of the KeyBindingTables key section in
TangramPackage.pkgdef.

The following Projects key section exists in TangramPackage.pkgdef in order to
register the new file template.

//The first GUID below is the GUID for the Miscellaneous Files project type, and can be changed
//  to the GUID of any other project you wish.
[$RootKey$\Projects\{A2FE74E1-B743-11d0-AE1A-00A0C90FFFC3}\AddItemTemplates\TemplateDirs\{19631222-1992-0612-1965-060119821989}\/1]
@="#100"
"TemplatesDir"="$PackageFolder$\Templates"
"SortPriority"=dword:00004E20

The contents of TangramPackage.vsdir, which is located a the location registered above are:

myext.myext|{ab02f9cb-42e8-467c-a242-d9bb2e1918a0}|#106|80|#109|{ab02f9cb-42e8-467c-a242-d9bb2e1918a0}|401|0|#107
The meaning of the fields are as follows:
	- Default.rtf - the default .RTF file
	- {ab02f9cb-42e8-467c-a242-d9bb2e1918a0} - same as CLSID_TangramPackagePackage
	- #106 - the literal value of IDS_EDITOR_NAME in TangramPackageUI.rc,
		which is displayed under the icon in the new file dialog.
	- 80 - the display ordering priority
	- #109 - the literal value of IDS_FILE_DESCRIPTION in TangramPackageUI.rc, which is displayed
		in the description window in the new file dialog.
	- {ab02f9cb-42e8-467c-a242-d9bb2e1918a0} - resource dll package guid
	- 401 - the literal value of IDI_FILE_ICON in TangramPackage.rc (not TangramPackageUI.rc), 
		which is the icon to display in the new file dialog.
	- 0 - template flags, which are unused here(we don't use this - see vsshell.idl)
	- #107 - the literal value of IDS_DEFAULT_NAME in TangramPackageUI.rc, which is the base
		name of the new files (i.e. myext1.myext, myext2.myext, etc.).

***************************************************************************/

class ATL_NO_VTABLE CAppDocPackage : 
	public CTangramPackageProxy,
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CAppDocPackage, &CLSID_TangramPackage>,
	public IVsPackageImpl<CAppDocPackage, &CLSID_TangramPackage>,
	public IOleCommandTargetImpl<CAppDocPackage>,
	public ATL::ISupportErrorInfoImpl<&__uuidof(IVsPackage)>
{
public:
	BEGIN_COM_MAP(CAppDocPackage)
		COM_INTERFACE_ENTRY(IVsPackage)
		COM_INTERFACE_ENTRY(IOleCommandTarget)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
	END_COM_MAP()

	VSL_DECLARE_NOT_COPYABLE(CAppDocPackage)

public:
	CAppDocPackage();	
	~CAppDocPackage();

	// This method will be called after IVsPackage::SetSite is called with a valid site
	void PostSited(IVsPackageEnums::SetSiteResult /*result*/);

	void PreClosing();

	void CreateTangramToolWnd();

	// This function provides the error information if it is not possible to load
	// the UI dll. It is for this reason that the resource IDS_E_BADINSTALL must
	// be defined inside this dll's resources.
	static const LoadUILibrary::ExtendedErrorInfo& GetLoadUILibraryErrorInfo()
	{
		static LoadUILibrary::ExtendedErrorInfo errorInfo(IDS_E_BADINSTALL);
		return errorInfo;
	}

	// DLL is registered with VS via a pkgdef file. Don't do anything if asked to
	// self-register.
	static HRESULT WINAPI UpdateRegistry(BOOL bRegister)
	{
		return S_OK;
	}

// NOTE - the arguments passed to these macros can not have names longer then 30 characters
// Definition of the commands handled by this package
VSL_BEGIN_COMMAND_MAP()
    VSL_COMMAND_MAP_ENTRY(CLSID_TangramPackageCmdSet, cmdidTangramCmd, NULL, CommandHandler::ExecHandler(&OnTangramCommand))
    VSL_COMMAND_MAP_ENTRY(CLSID_TangramPackageCmdSet, cmdidTangramTool, NULL, CommandHandler::ExecHandler(&OnTangramTool))
VSL_END_VSCOMMAND_MAP()


// The tool map implements IVsPackage::CreateTool that is called by VS to create a tool window 
// when appropriate.
VSL_BEGIN_TOOL_MAP()
    VSL_TOOL_ENTRY(CLSID_guidPersistanceSlot, m_ToolWindow.CreateAndShow())
VSL_END_TOOL_MAP()

// Command handler called when the user selects the command to show the toolwindow.
void OnTangramTool(CommandHandler* /*pSender*/, DWORD /*flags*/, VARIANT* /*pIn*/, VARIANT* /*pOut*/)
{
    m_ToolWindow.CreateAndShow();
}

// Command handler called when the user selects the "My Command" command.
void OnTangramCommand(CommandHandler* /*pSender*/, DWORD /*flags*/, VARIANT* /*pIn*/, VARIANT* /*pOut*/)
{
	m_ToolWindow.CreateAndShow();
}

void ShowMsg(CString strMsg)
{
	CComBSTR bstrTitle;
	VSL_CHECKBOOL_GLE(bstrTitle.LoadStringW(_AtlBaseModule.GetResourceInstance(), IDS_PROJNAME));
	// Get a pointer to the UI Shell service to show the message box.
	CComPtr<IVsUIShell> spUiShell = this->GetVsSiteCache().GetCachedService<IVsUIShell, SID_SVsUIShell>();
	LONG lResult;
	HRESULT hr = spUiShell->ShowMessageBox(
		0,
		CLSID_NULL,
		bstrTitle,
		CComBSTR(strMsg),
		NULL,
		0,
		OLEMSGBUTTON_OK,
		OLEMSGDEFBUTTON_FIRST,
		OLEMSGICON_INFO,
		0,
		&lResult);
	VSL_CHECKHRESULT(hr);
}


private:
    CAppXmlToolWnd m_ToolWindow;

	// Cookie returned when registering editor
	VSCOOKIE m_dwEditorCookie;
};

// This exposes CAppDocPackage for instantiation via DllGetClassObject; however, an instance
// can not be created by CoCreateInstance, as CAppDocPackage is specifically registered with
// VS, not the the system in general.
OBJECT_ENTRY_AUTO(CLSID_TangramPackage, CAppDocPackage)
