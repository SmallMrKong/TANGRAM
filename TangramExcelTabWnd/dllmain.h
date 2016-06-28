// dllmain.h : Declaration of module class.
#include "TangramExcelTabWnd_i.h"

class CTangramExcelTabWndApp : public CWinApp,
	public CAtlDllModuleT< CTangramExcelTabWndApp >
{
public:
	DECLARE_LIBID(LIBID_TangramExcelTabWndLib)
};

extern CTangramExcelTabWndApp theApp;
