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
#include "MSWord.h"
#include "WordPlusEvents.h"
using namespace Word;
using namespace OfficeCloudPlus::WordPlus::WordPlusEvent;

namespace OfficeCloudPlus
{
	namespace WordPlus
	{
		class CWordObject;
		class CWordCloudAddin;
		class CWordDocument : 
			public CWordDocumentEvents,
			public CTangramDocument,
			public map<HWND, CWordObject*>
		{
		public:
			CWordDocument(void);
			~CWordDocument(void);
			_Document*				m_pDoc;
			map<CString, CString>	m_mapDocUIInfo;

			void __stdcall OnClose();
			void InitDocument();
		};

		class CWordObject :
			public COfficeObject
		{
		public:
			CWordObject(void);

			BOOL					m_bDesignState;
			BOOL					m_bDesignTaskPane;

			CWordDocument*			m_pWordPlusDoc;
			void OnObjDestory();
		};
	}
}
