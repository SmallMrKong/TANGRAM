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

/*******************************************************************************
 * Copyright (c) 2006, 2009 IBM Corporation and others.
 * All rights reserved. This program and the accompanying materials
 * are made available under the terms of the Eclipse Public License v1.0
 * which accompanies this distribution, and is available at 
 * http://www.eclipse.org/legal/epl-v10.html
 * 
 * Contributors:
 *     IBM Corporation - initial API and implementation
 *******************************************************************************/
#ifndef ECLIPSE_JNI_H
#define ECLIPSE_JNI_H

#include "eclipseUnicode.h"
#include <jni.h>

typedef struct {
	int launchResult;
	int runResult;
	_TCHAR * errorMessage;
} JavaResults;

typedef jint (JNICALL *JNI_createJavaVM)(JavaVM **pvm, JNIEnv **env, void *args);

/* JNI Callback methods */
/* Use name mangling since we may be linking these from java with System.LoadLibrary */
#define set_exit_data 		Java_org_eclipse_equinox_launcher_JNIBridge__1set_1exit_1data
#define set_launcher_info	Java_org_eclipse_equinox_launcher_JNIBridge__1set_1launcher_1info
#define update_splash 		Java_org_eclipse_equinox_launcher_JNIBridge__1update_1splash
#define show_splash			Java_org_eclipse_equinox_launcher_JNIBridge__1show_1splash
#define get_splash_handle 	Java_org_eclipse_equinox_launcher_JNIBridge__1get_1splash_1handle
#define takedown_splash 	Java_org_eclipse_equinox_launcher_JNIBridge__1takedown_1splash
#define get_os_recommended_folder 	    Java_org_eclipse_equinox_launcher_JNIBridge__1get_1os_1recommended_1folder

//#ifdef __cplusplus
//extern "C" {
//#endif
///*
// * org_eclipse_equinox_launcher_JNIBridge#_set_exit_data
// * Signature: (Ljava/lang/String;Ljava/lang/String;)V
// */
//void set_exit_data(JNIEnv *, jobject, jstring, jstring);
//
///*
// * org_eclipse_equinox_launcher_JNIBridge#_set_launcher_info
// * Signature: (Ljava/lang/String;Ljava/lang/String;)V
// */
//void set_launcher_info(JNIEnv *, jobject, jstring, jstring);
//
///*
// * org_eclipse_equinox_launcher_JNIBridge#_update_splash
// * Signature: ()V
// */
//void update_splash(JNIEnv *, jobject);
//
///*
// * org_eclipse_equinox_launcher_JNIBridge#_get_splash_handle
// * Signature: ()J
// */
//jlong get_splash_handle(JNIEnv *, jobject);
//
///*
// * org_eclipse_equinox_launcher_JNIBridge#_show_splash
// * Signature: (Ljava/lang/String;)V
// */
//void show_splash(JNIEnv *, jobject, jstring);
//
///*
// * org_eclipse_equinox_launcher_JNIBridge#_takedown_splash
// * Signature: ()V
// */
//void takedown_splash(JNIEnv *, jobject);
//
///*
// * org_eclipse_equinox_launcher_JNIBridge#_get_os_recommended_folder
// * Signature: ()Ljava/lang/String
// */
//jstring get_os_recommended_folder(JNIEnv *, jobject);
//
//#ifdef __cplusplus
//}
//#endif

/* Start the Java VM and Wait For It to Terminate
 *
 * This method is responsible for starting the Java VM and for
 * detecting its termination. The resulting JVM exit code should
 * be returned to the main launcher, which will display a message if
 * the termination was not normal.
 */
extern JavaResults* startJavaJNI( _TCHAR* libPath, _TCHAR* vmArgs[], _TCHAR* progArgs[], _TCHAR* jarFile );

extern void cleanupVM( int );
#endif
