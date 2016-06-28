// EclipseHost.cpp 
//

#include "stdafx.h"
#include "EclipseHost.h"

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
	::CoInitialize(NULL);
	CComPtr<IDispatch> pDisp;
	pDisp.CoCreateInstance(L"Tangram.JavaProxy.1");
	return 0;
}

