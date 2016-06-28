/*******************************************************************************
 * Copyright (c) 2006, 2009 IBM Corporation and others.
 * All rights reserved. This program and the accompanying materials
 * are made available under the terms of the Eclipse Public License v1.0
 * which accompanies this distribution, and is available at 
 * http://www.eclipse.org/legal/epl-v10.html
 * 
 * Contributors:
 *     IBM Corporation - initial API and implementation
 * 	   Andrew Niefer
 *******************************************************************************/
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
 ********************************************************************************/

#include "stdafx.h"
#include "dllmain.h"
#include "eclipseJNI.h"
#include "eclipseCommon.h"
#include "eclipseOS.h"
#include "eclipseShm.h"
#include "TangramJavaProxy.h"

#include <shlobj.h>
#include <stdlib.h>
#include <string.h>


static _TCHAR* failedToLoadLibrary = _T_ECLIPSE("Failed to load the JNI shared library \"%s\".\n");
static _TCHAR* createVMSymbolNotFound = _T_ECLIPSE("The JVM shared library \"%s\"\ndoes not contain the JNI_CreateJavaVM symbol.\n");
static _TCHAR* failedCreateVM = _T_ECLIPSE("Failed to create the Java Virtual Machine.\n");
static _TCHAR* internalExpectedVMArgs = _T_ECLIPSE("Internal Error, the JVM argument list is empty.\n");
static _TCHAR* mainClassNotFound = _T_ECLIPSE("Failed to find a Main Class in \"%s\".\n");

/* local methods */
static jstring newJavaString(JNIEnv *env, _TCHAR * str);
static void registerNatives(JNIEnv *env);
static int shouldShutdown(JNIEnv *env);
static void JNI_ReleaseStringChars(JNIEnv *env, jstring s, const _TCHAR* data);
static const _TCHAR* JNI_GetStringChars(JNIEnv *env, jstring str);
static char * getMainClass(JNIEnv *env, _TCHAR * jarFile);
static void setLibraryLocation(JNIEnv *env, jobject obj);

static JavaVM * jvm = 0;
static JNIEnv *env = 0;

extern LPWSTR *szArglist;

/* cache String class and methods to avoid looking them up all the time */
static jclass string_class = NULL;
#if !defined(UNICODE) && !defined(MACOSX)
static jmethodID string_getBytesMethod = NULL;
static jmethodID string_ctor = NULL;
#endif

//////Begin Add by Tangram Team/////////////////////////
void InitCOMProxy()
{ 
	if (env&&_AtlModule.m_bInitCOM==false)
	{
		_AtlModule.m_bInitCOM = true;
		jclass TangramObj = env->FindClass("Tangram/COM/TangramObj");
		if (TangramObj != NULL)
		{
			ATLTRACE(_T("Init Com Proxy\n"));
			//int numNatives = sizeof(nativesCom) / sizeof(nativesCom[0]);
			//env->RegisterNatives(TangramObj, nativesCom, numNatives);
		}
		if (env->ExceptionOccurred() != 0) {
			env->ExceptionDescribe();
			env->ExceptionClear();
		}
	}
}
//////End Add by Tangram Team/////////////////////////

/* JNI Callback methods */
void set_exit_data(JNIEnv * env, jobject obj, jstring id, jstring s){
	const _TCHAR* data = NULL;
	const _TCHAR* sharedId = NULL;
	size_t length;
	 
	if(s != NULL) {
		length = env->GetStringLength(s);
		if(!env->ExceptionOccurred()) {
			data = JNI_GetStringChars(env, s);
			if (data != NULL) {
				if(id != NULL) {
					sharedId = JNI_GetStringChars(env, id);
					if(sharedId != NULL) {
						setSharedData(sharedId, data);
						JNI_ReleaseStringChars(env, id, sharedId);
					}
				} else {
					exitData = (_TCHAR*)malloc((length + 1) * sizeof(_TCHAR*));
					_tcsncpy( exitData, data, length);
					exitData[length] = _T_ECLIPSE('\0');
				}
				JNI_ReleaseStringChars(env, s, data);
			}
		}
		if(data == NULL && sharedId == NULL) {
			env->ExceptionDescribe();
			env->ExceptionClear();
		}
	}
}

void set_launcher_info(JNIEnv * env, jobject obj, jstring launcher, jstring name){
	const _TCHAR* launcherPath = NULL;
	const _TCHAR* launcherName = NULL;
	
	if (launcher != NULL) {
		launcherPath = JNI_GetStringChars(env, launcher);
		if (launcherPath != NULL) {
			setProgramPath(_tcsdup(launcherPath));
			JNI_ReleaseStringChars(env, launcher, launcherPath);
		}
	}
	
	if (name != NULL) {
		launcherName = JNI_GetStringChars(env, name);
		if (launcherName != NULL) {
			setOfficialName(_tcsdup(launcherName));
			JNI_ReleaseStringChars(env, name, launcherName);
		}
	}
}


void update_splash(JNIEnv * env, jobject obj){
	dispatchMessages();
}

jlong get_splash_handle(JNIEnv * env, jobject obj){
	return getSplashHandle();
}

void show_splash(JNIEnv * env, jobject obj, jstring s){
	const _TCHAR* data = NULL;
	
	setLibraryLocation(env, obj);
	
	if(s != NULL) {
		data = JNI_GetStringChars(env, s);
		if(data != NULL) {
			showSplash(data);
			JNI_ReleaseStringChars(env, s, data);
		} else {
			env->ExceptionDescribe();
			env->ExceptionClear();
		}
	}
}

void  takedown_splash(JNIEnv * env, jobject obj){
	takeDownSplash();
}

jstring get_os_recommended_folder(JNIEnv * env, jobject obj){
#ifdef MACOSX
	return newJavaString(env, getFolderForApplicationData());
#else
	return NULL;
#endif
}

//static JNINativeMethod natives[] = 
//{
//	{ "_update_splash", "()V", (void *)&update_splash },
//	{ "_get_splash_handle", "()J", (void *)&get_splash_handle },
//	{ "_set_exit_data", "(Ljava/lang/String;Ljava/lang/String;)V", (void *)&set_exit_data },
//	{ "_set_launcher_info", "(Ljava/lang/String;Ljava/lang/String;)V", (void *)&set_launcher_info },
//	{ "_show_splash", "(Ljava/lang/String;)V", (void *)&show_splash },
//	{ "_takedown_splash", "()V", (void *)&takedown_splash },
//	{ "_get_os_recommended_folder", "()Ljava/lang/String;", (void *)&get_os_recommended_folder } 
//};

/*
 * On AIX we need the location of the eclipse shared library so that we
 * can find the libeclipse-motif.so library.  Reach into the JNIBridge
 * object to get the "library" field.
 */
void setLibraryLocation(JNIEnv * env, jobject obj) {
	jclass bridge = env->FindClass("org/eclipse/equinox/launcher/JNIBridge");
	if (bridge != NULL) {
		jfieldID libraryField = env->GetFieldID(bridge, "library", "Ljava/lang/String;");
		if (libraryField != NULL) {
			jstring stringObject = (jstring) env->GetObjectField(obj, libraryField);
			if (stringObject != NULL) 
			{
				const _TCHAR * str = JNI_GetStringChars(env, stringObject);
				eclipseLibrary = _tcsdup(str);
				JNI_ReleaseStringChars(env, stringObject, str);
			}
		}
	}
	if( env->ExceptionOccurred() != 0 ){
		env->ExceptionDescribe();
		env->ExceptionClear();
	}
}

void registerNatives(JNIEnv *env) {
	///*begin Add by Tangram Team*/
	//if (m_pTangram)
	//{
	//	m_pTangram->put_AppKeyValue(CComBSTR(L"JNIEnv"), CComVariant((long)env));
	//}
	///*end Add by Tangram Team*/
	jclass bridge = env->FindClass("org/eclipse/equinox/launcher/JNIBridge");
	if(bridge != NULL) {
		JNINativeMethod natives[] =
		{
			{ "_update_splash", "()V", (void *)&update_splash },
			{ "_get_splash_handle", "()J", (void *)&get_splash_handle },
			{ "_set_exit_data", "(Ljava/lang/String;Ljava/lang/String;)V", (void *)&set_exit_data },
			{ "_set_launcher_info", "(Ljava/lang/String;Ljava/lang/String;)V", (void *)&set_launcher_info },
			{ "_show_splash", "(Ljava/lang/String;)V", (void *)&show_splash },
			{ "_takedown_splash", "()V", (void *)&takedown_splash },
			{ "_get_os_recommended_folder", "()Ljava/lang/String;", (void *)&get_os_recommended_folder }
		};
		int numNatives = sizeof(natives) / sizeof(natives[0]);
		env->RegisterNatives(bridge, natives, numNatives);	
	}
	if( env->ExceptionOccurred() != 0 ){
		env->ExceptionDescribe();
		env->ExceptionClear();
	}
}


/* Get a _TCHAR* from a jstring, string should be released later with JNI_ReleaseStringChars */
static const _TCHAR * JNI_GetStringChars(JNIEnv *env, jstring str) {
	const _TCHAR * result = NULL;
#ifdef UNICODE
	/* GetStringChars is not null terminated, make a copy */
	const _TCHAR * stringChars = (_TCHAR *)env->GetStringChars(str, 0);
	int length = env->GetStringLength(str);
	_TCHAR * copy = (_TCHAR *)malloc( (length + 1) * sizeof(_TCHAR));
	_tcsncpy(copy, stringChars, length);
	copy[length] = _T_ECLIPSE('\0');
	env->ReleaseStringChars(str, (const jchar *)stringChars);
	result = copy;
#elif MACOSX
	/* Use UTF on the Mac */
	result = env->GetStringUTFChars(env, str, 0);
#else
	/* Other platforms, use java's default encoding */ 
	_TCHAR* buffer = NULL;
	if (string_class == NULL)
		string_class = env->FindClass(env, "java/lang/String");
	if (string_class != NULL) {
		if (string_getBytesMethod == NULL)
			string_getBytesMethod = env->GetMethodID(env, string_class, "getBytes", "()[B");
		if (string_getBytesMethod != NULL) {
			jbyteArray bytes = env->CallObjectMethod(env, str, string_getBytesMethod);
			if (!env->ExceptionOccurred()) {
				jsize length = env->GetArrayLength(env, bytes);
				buffer = malloc( (length + 1) * sizeof(_TCHAR*));
				env->GetByteArrayRegion(env, bytes, 0, length, (jbyte*)buffer);
				buffer[length] = 0;
			}
			env->DeleteLocalRef(env, bytes);
		}
	}
	if(buffer == NULL) {
		env->ExceptionDescribe();
		env->ExceptionClear();
	}
	result = buffer;
#endif
	return result;
}

/* Release the string that was obtained using JNI_GetStringChars */
static void JNI_ReleaseStringChars(JNIEnv *env, jstring s, const _TCHAR* data) {
#ifdef UNICODE
	free((_TCHAR*)data);
#elif MACOSX
	env->ReleaseStringUTFChars(env, s, data);
#else
	free((_TCHAR*)data);
#endif
}

static jstring newJavaString(JNIEnv *env, _TCHAR * str)
{
	jstring newString = NULL;
#ifdef UNICODE
	size_t length = _tcslen(str);
	newString = env->NewString((const jchar*)str, length);
#elif MACOSX
	newString = env->NewStringUTF(env, str);
#else
	size_t length = _tcslen(str);
	jbyteArray bytes = env->NewByteArray(env, length);
	if(bytes != NULL) {
		env->SetByteArrayRegion(env, bytes, 0, length, (jbyte *)str);
		if (!env->ExceptionOccurred()) {
			if (string_class == NULL)
				string_class = env->FindClass(env, "java/lang/String");
			if(string_class != NULL) {
				if (string_ctor == NULL)
					string_ctor = env->GetMethodID(env, string_class, "<init>",  "([B)V");
				if(string_ctor != NULL) {
					newString = env->NewObject(env, string_class, string_ctor, bytes);
				}
			}
		}
		env->DeleteLocalRef(env, bytes);
	}
#endif
	if(newString == NULL) {
		env->ExceptionDescribe();
		env->ExceptionClear();
	}
	return newString;
}

static jobjectArray createRunArgs( JNIEnv *env, _TCHAR * args[] ) {
	int index = 0, length = -1;
	jobjectArray stringArray = NULL;
	jstring string;
	
	/*count the number of elements first*/
	while(args[++length] != NULL);
	
	if (string_class == NULL)
		string_class = env->FindClass("java/lang/String");
	if(string_class != NULL) {
		stringArray = env->NewObjectArray(length, string_class, 0);
		if(stringArray != NULL) {
			for( index = 0; index < length; index++) {
				string = newJavaString(env, args[index]);
				if(string != NULL) {
					env->SetObjectArrayElement(stringArray, index, string); 
					env->DeleteLocalRef(string);
				} else {
					env->DeleteLocalRef(stringArray);
					env->ExceptionDescribe();
					env->ExceptionClear();
					return NULL;
				}
			}
		}
	} 
	if(stringArray == NULL) {
		env->ExceptionDescribe();
		env->ExceptionClear();
	}
	return stringArray;
}
					 
JavaResults * startJavaJNI( _TCHAR* libPath, _TCHAR* vmArgs[], _TCHAR* progArgs[], _TCHAR* jarFile )
{
	int i;
	int numVMArgs = -1;
	void * jniLibrary;
	JNI_createJavaVM createJavaVM;
	JavaVMInitArgs init_args;
	JavaVMOption * options;
	char * mainClassName = NULL;
	JavaResults * results = NULL;
	
	/* JNI reflection */
	jclass mainClass = NULL;			/* The Main class to load */
	jmethodID mainConstructor = NULL;	/* Main's default constructor Main() */
	jobject mainObject = NULL;			/* An instantiation of the main class */
	jmethodID runMethod = NULL;			/* Main.run(String[]) */
	jobjectArray methodArgs = NULL;		/* Arguments to pass to run */
	
	results = (JavaResults *)malloc(sizeof(JavaResults));
	memset(results, 0, sizeof(JavaResults));
	
	jniLibrary = loadLibrary(libPath);
	if(jniLibrary == NULL) {
		results->launchResult = -1;
		results->errorMessage = (_TCHAR *)malloc((_tcslen(failedToLoadLibrary) + _tcslen(libPath) + 1) * sizeof(_TCHAR));
		_stprintf(results->errorMessage, failedToLoadLibrary, libPath);
		return results; /*error*/
	}

	createJavaVM = (JNI_createJavaVM)findSymbol(jniLibrary, _T_ECLIPSE("JNI_CreateJavaVM"));
	if(createJavaVM == NULL) {
		results->launchResult = -2;
		results->errorMessage = (_TCHAR *)malloc((_tcslen(createVMSymbolNotFound) + _tcslen(libPath) + 1) * sizeof(_TCHAR));
		_stprintf(results->errorMessage, createVMSymbolNotFound, libPath);
		return results; /*error*/
	}
	
	/* count the vm args */
	while(vmArgs[++numVMArgs] != NULL) {}
	
	if(numVMArgs <= 0) {
		/*error, we expect at least the required vm arg */
		results->launchResult = -3;
		results->errorMessage = _tcsdup(internalExpectedVMArgs);
		return results;
	}
	///begin Add by Tangram Team/////
	bool bAppendClassPath = false;
	///end Add by Tangram Team/////
	options = (JavaVMOption *)malloc(numVMArgs * sizeof(JavaVMOption));
	for(i = 0; i < numVMArgs; i++){
		char* str = toNarrow(vmArgs[i]);
		///begin Add by Tangram Team/////
		//if (bAppendClassPath==false&&m_pTangram)
		//{
		//	CStringA stroption = CStringA(str);
		//	if (stroption.Find("java.class.path=") != -1)
		//	{
		//		bAppendClassPath = true;
		//		//TCHAR									m_szBuffer[MAX_PATH];
		//		//SHGetFolderPath(NULL, CSIDL_PROGRAM_FILES, NULL, 0, m_szBuffer);
		//		int nPos = stroption.ReverseFind('/');
		//		if(nPos==-1)
		//			nPos = stroption.ReverseFind('\\');
		//		CStringA str2 = stroption.Left(nPos) + ("/tangram.jar");
		//		//CStringA str2 = CStringA(m_szBuffer);
		//		//str2+="//tangram//tangram.jar";
		//		nPos = str2.Find("=");
		//		str2 = str2.Mid(nPos + 1);
		//		//str2.Replace("\\", "//");
		//		//str2.Replace("\\\\", "//");
		//		stroption = stroption +";" + str2;

		//		//GetModuleFileName(GetModuleHandle(L"Tangram.dll"), szPath, MAX_PATH);
		//		//CStringA strPath = stroption + ";" + CStringA(szPath);
		//		//strPath = strPath.Left(nPos);
		//		//strPath += _T(".jar");
		//		//strPath.Replace(" ", "%20");
		//		str = (char*)LPCSTR(stroption);
		//		//strPath.ReleaseBuffer();
		//	}
		//}
		///end Add by Tangram Team/////

		options[i].optionString = str;
		options[i].extraInfo = 0;
	}
		
#ifdef MACOSX
	init_args.version = JNI_VERSION_1_4;
#else		
	init_args.version = JNI_VERSION_1_2;
#endif
	init_args.options = options;
	init_args.nOptions = numVMArgs;
	init_args.ignoreUnrecognized = JNI_TRUE;
	
	if( createJavaVM(&jvm, &env, &init_args) == 0 ) {
		/*begin Add by Tangram Team*/
		if (m_pTangram)
		{
			m_pTangram->put_AppKeyValue(CComBSTR(L"JVM"), CComVariant((long)jvm));
		}
		/*end Add by Tangram Team*/
		registerNatives(env);
		
		mainClassName = getMainClass(env, jarFile);
		if (mainClassName != NULL) {
			mainClass = env->FindClass(mainClassName);
			free(mainClassName);
		}
		
		if (mainClass == NULL) {
			if (env->ExceptionOccurred()) {
				env->ExceptionDescribe();
				env->ExceptionClear();
			}
			mainClass = env->FindClass("org/eclipse/equinox/launcher/Main");
		}	

		if(mainClass != NULL) {
			results->launchResult = -6; /* this will be reset to 0 below on success */
			mainConstructor = env->GetMethodID(mainClass, "<init>", "()V");
			if(mainConstructor != NULL) {
				mainObject = env->NewObject(mainClass, mainConstructor);
				if(mainObject != NULL) {
					runMethod = env->GetMethodID(mainClass, "run", "([Ljava/lang/String;)I");
					if(runMethod != NULL) {
						methodArgs = createRunArgs(env, progArgs);
						if(methodArgs != NULL) {
							results->launchResult = 0;
							results->runResult = env->CallIntMethod(mainObject, runMethod, methodArgs);
							env->DeleteLocalRef(methodArgs);
						}
					}
					env->DeleteLocalRef(mainObject);
				}
			}
		} else {
			results->launchResult = -5;
			results->errorMessage = (_TCHAR *)malloc((_tcslen(mainClassNotFound) + _tcslen(jarFile) + 1) * sizeof(_TCHAR));
			_stprintf(results->errorMessage, mainClassNotFound, jarFile);
		}
		if(env->ExceptionOccurred()){
			env->ExceptionDescribe();
			env->ExceptionClear();
		}
		
	} else {
		results->launchResult = -4;
		results->errorMessage = _tcsdup(failedCreateVM);
	}

	/* toNarrow allocated new strings, free them */
	//for(i = 0; i < numVMArgs; i++){
	//	free( options[i].optionString );
	//}
	free(options);
	return results;
}

static char * getMainClass(JNIEnv *env, _TCHAR * jarFile) 
{
	jclass jarFileClass = NULL, manifestClass = NULL, attributesClass = NULL;
	jmethodID jarFileConstructor = NULL, getManifestMethod = NULL, getMainAttributesMethod = NULL, closeJarMethod = NULL, getValueMethod = NULL;
	jobject jarFileObject, manifest, attributes;
	jstring mainClassString = NULL;
	jstring jarFileString, headerString;
	const _TCHAR *mainClass;
	
	/* get the classes we need */
	jarFileClass = env->FindClass("java/util/jar/JarFile");
	if (jarFileClass != NULL) {
		manifestClass = env->FindClass("java/util/jar/Manifest");
		if (manifestClass != NULL) {
			attributesClass = env->FindClass("java/util/jar/Attributes");
		}
	}
	if (env->ExceptionOccurred()) {
		env->ExceptionDescribe();
		env->ExceptionClear();
	}
	if (attributesClass == NULL)
		return NULL;
	
	/* find the methods */
	jarFileConstructor = env->GetMethodID(jarFileClass, "<init>", "(Ljava/lang/String;Z)V");
	if(jarFileConstructor != NULL) {
		getManifestMethod = env->GetMethodID(jarFileClass, "getManifest", "()Ljava/util/jar/Manifest;");
		if(getManifestMethod != NULL) {
			closeJarMethod = env->GetMethodID(jarFileClass, "close", "()V");
			if (closeJarMethod != NULL) {
				getMainAttributesMethod = env->GetMethodID(manifestClass, "getMainAttributes", "()Ljava/util/jar/Attributes;");
				if (getMainAttributesMethod != NULL) {
					getValueMethod = env->GetMethodID(attributesClass, "getValue", "(Ljava/lang/String;)Ljava/lang/String;");
				}
			}
		}
	}
	if (env->ExceptionOccurred()) {
		env->ExceptionDescribe();
		env->ExceptionClear();
	}
	if (getValueMethod == NULL)
		return NULL;
	
	/* jarFileString = new String(jarFile); */
	jarFileString = newJavaString(env, jarFile);
	 /* headerString = new String("Main-Class"); */
	 headerString = newJavaString(env, _T_ECLIPSE("Main-Class"));
	if (jarFileString != NULL && headerString != NULL) {
		/* jarfileObject = new JarFile(jarFileString, false); */
		jarFileObject = env->NewObject(jarFileClass, jarFileConstructor, jarFileString, JNI_FALSE);
		if (jarFileObject != NULL) { 
			/* manifest = jarFileObject.getManifest(); */
			 manifest = env->CallObjectMethod(jarFileObject, getManifestMethod);
			 if (manifest != NULL) {
				 /*jarFileObject.close() */
				 env->CallVoidMethod(jarFileObject, closeJarMethod);
				 if (!env->ExceptionOccurred()) {
					 /* attributes = manifest.getMainAttributes(); */
					 attributes = env->CallObjectMethod(manifest, getMainAttributesMethod);
					 if (attributes != NULL) {
						 /* mainClassString = attributes.getValue(headerString); */
						 mainClassString = (jstring)env->CallObjectMethod(attributes, getValueMethod, headerString);
					 }
				 }
			 }
			 env->DeleteLocalRef(jarFileObject);
		}
	}
	
	if (jarFileString != NULL)
		env->DeleteLocalRef(jarFileString);
	if (headerString != NULL)
		env->DeleteLocalRef(headerString);
	
	if (env->ExceptionOccurred()) {
		env->ExceptionDescribe();
		env->ExceptionClear();
	}
	
	if (mainClassString == NULL)
		return NULL;
	
	mainClass = JNI_GetStringChars(env, mainClassString);
	if(mainClass != NULL) {
		int i = -1;
		char *result = toNarrow(mainClass);
		JNI_ReleaseStringChars(env, mainClassString, mainClass);
		
		/* replace all the '.' with '/' */
		while(result[++i] != '\0') {
			if(result[i] == '.')
				result[i] = '/';
		}
		return result;
	}
	return NULL;
}

void cleanupVM(int exitCode) {
	JNIEnv * localEnv = env;
	if (jvm == 0)
		return;
	
	if (secondThread)
		jvm->AttachCurrentThread((void**)&localEnv, NULL);
	else
		localEnv = env;
	if (localEnv == 0)
		return;
	
	/* we call System.exit() unless osgi.noShutdown is set */
	if (shouldShutdown(env)) {
		jclass systemClass = NULL;
		jmethodID exitMethod = NULL;
		systemClass = env->FindClass("java/lang/System");
		try
		{
			if (systemClass != NULL) {
				exitMethod = env->GetStaticMethodID(systemClass, "exit", "(I)V");
				if (exitMethod != NULL) {
					if (_AtlModule.m_pJavaProxy)
					{
						DWORD dw = _AtlModule.m_pJavaProxy->Release();
						while(dw)
							dw = _AtlModule.m_pJavaProxy->Release();
					}
					if (szArglist) {
						//LocalFree(szArglist);
					}
					env->CallStaticVoidMethod(systemClass, exitMethod, exitCode);
				}
			}
		}
		catch(...)
		{
			int i = 0;
		}
		if (env->ExceptionOccurred()) {
			env->ExceptionDescribe();
			env->ExceptionClear();
		}
	}
	jvm->DestroyJavaVM();
}

static int shouldShutdown(JNIEnv * env) {
	jclass booleanClass = NULL;
	jmethodID method = NULL;
	jstring arg = NULL;
	jboolean result = 0;
	
	booleanClass = env->FindClass("java/lang/Boolean");
	if (booleanClass != NULL) {
		method = env->GetStaticMethodID(booleanClass, "getBoolean", "(Ljava/lang/String;)Z");
		if (method != NULL) {
			arg = newJavaString(env, _T_ECLIPSE("osgi.noShutdown"));
			result = env->CallStaticBooleanMethod(booleanClass, method, arg);
			env->DeleteLocalRef(arg);
		}
	}
	if (env->ExceptionOccurred()) {
		env->ExceptionDescribe();
		env->ExceptionClear();
	}
	return (result == 0);
}


