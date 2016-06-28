// TangramJavaProxy.cpp : CTangramJavaProxy µÄÊµÏÖ

#include "stdafx.h"
#include "dllmain.h"
#include "TangramJavaProxy.h"
#include "TangramEclipse_i.c"
#include "tangram.c"
#include <shellapi.h>
#include <shlobj.h>
#define MAX_LOADSTRING 100
#ifdef _WIN32
#include <direct.h>
#else
#include <unistd.h>
#endif
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <ctype.h>
#include <locale.h>
#include <sys/stat.h>
#include "eclipseUnicode.h"
#include "eclipseJni.h"
#include "eclipseCommon.h"

#define NAME         _T("-name")
#define VMARGS       _T("-vmargs")					/* special option processing required */
#define LIBRARY		  _T("--launcher.library")
#define SUPRESSERRORS _T("--launcher.suppressErrors")
#define INI			  _T("--launcher.ini")
#define PROTECT 	 _T("-protect")
#define PROTECT	      _T("-protect")	/* This argument is also handled in eclipse.c for Mac specific processing */
#define ROOT		  _T("root")		/* the only level of protection we care now */

//_TCHAR* libraryMsg = _T("The %s executable launcher was unable to locate its \ncompanion shared library.");
//_TCHAR* entryMsg = _T("There was a problem loading the shared library and \nfinding the entry point.");

_TCHAR*  name = NULL;			/* program name */
_TCHAR** userVMarg = NULL;     		/* user specific args for the Java VM */
_TCHAR*  officialName = NULL;
int      suppressErrors;/* = 0;	*/			/* supress error dialogs */
int      protectRoot = 0;				/* check if launcher was run as root, currently works only on Linux/UNIX platforms */
 
LPWSTR *szArglist = nullptr;

extern int initialArgc;
extern _TCHAR** initialArgv;
extern _TCHAR* eclipseLibrary;// = NULL; /* path to the eclipse shared library */

void setInitialArgsW(int argc, _TCHAR** argv, _TCHAR* lib);
int runW(int argc, _TCHAR* argv[], _TCHAR* vmArgs[]);
int readIniFile(_TCHAR* program, int *argc, _TCHAR ***argv);
int readConfigFile(_TCHAR * config_file, int *argc, _TCHAR ***argv);
_TCHAR* getIniFile(_TCHAR* program, int consoleLauncher);

_TCHAR* findProgram(_TCHAR* argv[]) {
	_TCHAR * program;
	/* windows, make sure we are looking for the .exe */
	_TCHAR * ch;
	int length = _tcslen(argv[0]);
	ch = (_TCHAR*)malloc((length + 5) * sizeof(_TCHAR));
	_tcscpy(ch, argv[0]);

	if (length <= 4 || _tcsicmp(&ch[length - 4], _T(".exe")) != 0)
		_tcscat(ch, _T_ECLIPSE(".exe"));

	program = findCommand(ch);
	if (ch != program)
		free(ch);
	if (program == NULL)
	{
		program = (_TCHAR*)malloc(MAX_PATH_LENGTH + 1);
		GetModuleFileName(NULL, program, MAX_PATH_LENGTH);
		argv[0] = program;
	}
	else if (_tcscmp(argv[0], program) != 0) {
		argv[0] = program;
	}
	return program;
}

/*
* Parse arguments of the command.
*/
void parseArgs(int* pArgc, _TCHAR* argv[], int useVMargs)
{
	int     index;

	/* Ensure the list of user argument is NULL terminated. */
	argv[*pArgc] = NULL;

	/* For each user defined argument */
	for (index = 0; index < *pArgc; index++) {
		if (_tcsicmp(argv[index], VMARGS) == 0) {
			if (useVMargs == 1) { //Use the VMargs as the user specified vmArgs
				userVMarg = &argv[index + 1];
			}
			argv[index] = NULL;
			*pArgc = index;
		}
		else if (_tcsicmp(argv[index], NAME) == 0) {
			name = argv[++index];
		}
		else if (_tcsicmp(argv[index], LIBRARY) == 0) {
			//eclipseLibrary = argv[++index];
			index++;
		}
		else if (_tcsicmp(argv[index], SUPRESSERRORS) == 0) {
			suppressErrors = 1;
		}
		else if (_tcsicmp(argv[index], PROTECT) == 0) {
			if (_tcsicmp(argv[++index], ROOT) == 0) {
				protectRoot = 1;
			}
		}
	}
}

/* We need to look for --launcher.ini before parsing the other args */
_TCHAR* checkForIni(int argc, _TCHAR* argv[])
{
	int index;
	for (index = 0; index < (argc - 1); index++) {
		if (_tcsicmp(argv[index], INI) == 0) {
			return argv[++index];
		}
	}
	return NULL;
}

/*
* Create a new array containing user arguments from the config file first and
* from the command line second.
* Allocate an array large enough to host all the strings passed in from
* the argument configArgv and argv. That array is passed back to the
* argv argument. That array must be freed with the regular free().
* Note that both arg lists are expected to contain the argument 0 from the C
* main method. That argument contains the path/executable name. It is
* only copied once in the resulting list.
*
* Returns 0 if success.
*/
int createUserArgs(int configArgc, _TCHAR **configArgv, int *argc, _TCHAR ***argv)
{
	_TCHAR** newArray = (_TCHAR **)malloc((configArgc + *argc + 1) * sizeof(_TCHAR *));

	newArray[0] = (*argv)[0];	/* use the original argv[0] */
	memcpy(newArray + 1, configArgv, configArgc * sizeof(_TCHAR *));

	/* Skip the argument zero (program path and name) */
	memcpy(newArray + 1 + configArgc, *argv + 1, (*argc - 1) * sizeof(_TCHAR *));

	/* Null terminate the new list of arguments and return it. */
	*argv = newArray;
	*argc += configArgc;
	(*argv)[*argc] = NULL;

	return 0;
}

/*
* Determine the default official application name
*
* This function provides the default application name that appears in a variety of
* places such as: title of message dialog, title of splash screen window
* that shows up in Windows task bar.
* It is computed from the name of the launcher executable and
* by capitalizing the first letter. e.g. "c:/ide/eclipse.exe" provides
* a default name of "Eclipse".
*/
_TCHAR* getDefaultOfficialName(_TCHAR* program)
{
	_TCHAR *ch = NULL;

	/* Skip the directory part */
	ch = lastDirSeparator(program);
	if (ch == NULL) ch = program;
	else ch++;

	ch = _tcsdup(ch);
#ifdef _WIN32
	{
		/* Search for the extension .exe and cut it */
		_TCHAR *extension = _tcsrchr(ch, _T_ECLIPSE('.'));
		if (extension != NULL)
		{
			*extension = _T_ECLIPSE('\0');
		}
	}
#endif
	/* Upper case the first character */
#ifndef LINUX
	{
		*ch = _totupper(*ch);
	}
#else
	{
		if (*ch >= 'a' && *ch <= 'z')
		{
			*ch -= 32;
		}
	}
#endif
	return ch;
}

// CTangramJavaProxy

CTangramJavaProxy::CTangramJavaProxy()
{
	_AtlModule.m_pJavaProxy = this;
}

CTangramJavaProxy::~CTangramJavaProxy()
{
	ATLTRACE(_T("delete CTangramJavaProxy: %x\n"), this);
}

HRESULT CTangramJavaProxy::FinalConstruct()
{
	if (m_pTangram == nullptr)
		InitEclipse();
	return S_OK;
}

void CTangramJavaProxy::FinalRelease()
{
}

STDMETHODIMP CTangramJavaProxy::InitEclipse()
{
	////////begin Add by Tangram Team////////////////////////
	CComPtr<ITangram> pDisp;
	pDisp.CoCreateInstance(L"Tangram.tangram.1");
	if (pDisp)
	{
		m_pTangram = pDisp.Detach();
		m_pTangram->put_JavaProxy((IJavaProxy*)this);
	}
	_TCHAR** configArgv = nullptr;
	int 	 configArgc = 0;
	int      ret = 0;
	int		nArgs;
	
	szArglist = CommandLineToArgvW(GetCommandLineW(), &nArgs);

	TCHAR	m_szBuffer[MAX_PATH];
	::GetModuleFileName(::GetModuleHandle(L"tangrameclipse.dll"), m_szBuffer, MAX_PATH);
	eclipseLibrary = m_szBuffer;
	////////end Add by Tangram Team////////////////////////

	setlocale(LC_ALL, "");

	initialArgc = nArgs;
	initialArgv = (_TCHAR**)malloc((nArgs + 1) * sizeof(_TCHAR*));
	memcpy(initialArgv, szArglist, (nArgs + 1) * sizeof(_TCHAR*));

	/*
	* Strip off any extroneous <CR> from the last argument. If a shell script
	* on Linux is created in DOS format (lines end with <CR><LF>), the C-shell
	* does not strip off the <CR> and hence the argument is bogus and may
	* not be recognized by the launcher or eclipse itself.
	*/
	_TCHAR*  ch = _tcschr(szArglist[nArgs - 1], _T('\r'));
	if (ch != NULL)
	{
		*ch = _T('\0');
	}

	/* Determine the full pathname of this program. */
	_TCHAR*  program = findProgram(szArglist);

	/* Parse configuration file arguments */
	_TCHAR*  iniFile = checkForIni(nArgs, szArglist);
	if (iniFile != NULL)
		ret = readConfigFile(iniFile, &configArgc, &configArgv);
	else
		ret = readIniFile(program, &configArgc, &configArgv);
	if (ret == 0)
	{
		parseArgs(&configArgc, configArgv, 0);
	}

	/* Special case - user arguments specified in the config file
	* are appended to the user arguments passed from the command line.
	*/
	if (configArgc > 0)
	{
		createUserArgs(configArgc, configArgv, &nArgs, &szArglist);
	}

	///* Initialize official program name */
	//officialName = name != NULL ? _tcsdup(name) : getDefaultOfficialName(program);
	//free(officialName);

	free(program);
	runW(nArgs, szArglist, userVMarg);
	LocalFree(szArglist);

	return S_OK;
}

STDMETHODIMP CTangramJavaProxy::InitComProxy()
{
	InitCOMProxy();
	return S_OK;
}
