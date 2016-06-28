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
* OUTLINED IN THE TANGRAM LICENSE AGREEMENT.TANGRAM TEAM
* GRANTS TO YOU (ONE SOFTWARE DEVELOPER) THE LIMITED RIGHT TO USE
* THIS SOFTWARE ON A SINGLE COMPUTER.
*
* CONTACT INFORMATION:
* mailto:sunhuizlz@yeah.net
* http://www.CloudAddin.com
*
********************************************************************************/


#pragma once

#include "excel.h"
#include "../CloudAddin.h"
#include "ExcelPlusEvents.h"

using namespace Excel;
using namespace OfficeCloudPlus::ExcelPlus::ExcelPlusEvent;

namespace OfficeCloudPlus
{
	namespace ExcelPlus
	{
		class CExcelObject;
		class CExcelWorkSheet;
		class CExcelWorkBook;

		class CExcelCloudAddin :
			public CCloudAddin,
			public map<_Workbook*, CExcelWorkBook*>
		{
		public:
			CExcelCloudAddin();
			virtual ~CExcelCloudAddin();

			CExcelObject*					m_pActiveExcelObject;
			CComPtr<Excel::_Application>	m_pExcelApplication;

			STDMETHOD(AddDocXml)(IDispatch* pDocdisp, BSTR bstrXml, BSTR bstrKey, VARIANT_BOOL* bSuccess);
			STDMETHOD(GetDocXmlByKey)(IDispatch* pDocdisp, BSTR bstrKey, BSTR* pbstrXml);
			STDMETHOD(CreateOfficeDocument)(BSTR bstrXml);
			STDMETHOD(ExportOfficeObjXml)(IDispatch* OfficeObject, BSTR* bstrXml);

		private:
			bool							m_bOldVer;

			void OnOpenDoc(WPARAM);
			void SetFocus(HWND);
			void ProcessMsg(LPMSG lpMsg);
			//void OnDestroyVbaForm(HWND);
			void OnVbaFormInit(HWND, CWndFrame*);
			void OnWorkBookActivate(CExcelObject* pExcelObject);
			void UpdateOfficeObj(IDispatch* pObj, CString strXml, CString strName);
			CString GetFormXml(CString strFormName);
			//CTangramApp:
			void WindowCreated(LPCTSTR strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd);
			void WindowDestroy(HWND hWnd);
			STDMETHOD(TangramCommand)(IDispatch* RibbonControl);
			HRESULT OnConnection(IDispatch* pHostApp, int ConnectMode);
			HRESULT OnDisconnection(int DisConnectMode)
			{
				return S_OK;
			};

			STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);
			//void InsertComponent(CExcelWorkBookWnd* pWnd, CString strXml);
			void ConnectOfficeObj(HWND hWnd);
			bool OnActiveOfficeObj(HWND hWnd);
		};

		class CExcelObject :
			public COfficeObject
		{
		public:
			CExcelObject(void);
			~CExcelObject(void);

			HWND				m_hExcelEdit;
			HWND				m_hExcelEdit2;

			CString				m_strActiveSheetName;

			CExcelWorkBook*	    m_pWorkBook;
			void ProcessMouseDownMsg(HWND);
			void OnObjDestory();
			void SheetNameChanged();
			void OnOpenDoc();
		};
	}
}


