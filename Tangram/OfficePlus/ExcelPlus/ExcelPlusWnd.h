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
#include "../CloudAddin.h"
#include "ExcelPlusEvents.h"
using namespace OfficeCloudPlus;
using namespace OfficeCloudPlus::ExcelPlus::ExcelPlusEvent;

namespace OfficeCloudPlus
{
	namespace ExcelPlus
	{
		class CExcelObject;
		class CExcelCloudAddin;
		class CExcelWorkBook;
		class CExcelWorkSheet;
		typedef map<HWND, CExcelObject*> ExcelWorkBookWndMap;
		typedef ExcelWorkBookWndMap::iterator ExcelWorkBookWndMapIT;

		class CExcelWorkBook : 
			public CTangramDocument,
			public CExcelWorkbookEvents,
			public map<_Worksheet*, CExcelWorkSheet*>
		{
		public:
			CExcelWorkBook(void);
			~CExcelWorkBook(void);


			CString						m_strDefaultSheetXml;
			_Workbook*					m_pWorkBook;
			ExcelWorkBookWndMap			m_mapExcelWorkBookWnd;

			CWndNode*					m_pSheetNode;
			map<CString, CString>		m_mapWorkSheetInfo;

			void InitWorkBook();
			void ModifySheetForTangram(IDispatch* Sh, CString strSheetXml, CString strTaskPaneXml);
			//void AddPropertyToSheet()
		private:
			void __stdcall OnNewSheet(IDispatch* Sh);
			void __stdcall OnSheetActivate(IDispatch* Sh);
			void __stdcall OnSheetBeforeDelete(IDispatch* Sh);
			void __stdcall OnBeforeSave(VARIANT_BOOL SaveAsUI, VARIANT_BOOL* Cancel);
		};

		class CExcelWorkSheet : public map<CString, CString>
		{
		public:
			CExcelWorkSheet(void);
			~CExcelWorkSheet(void);
			CString m_strKey;
			CString m_strSheetName;
			_Worksheet* m_pSheet;
			map<CString, CWndNode*> m_mapNodeMap;
		};
	}
}
