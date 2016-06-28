

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Sun Jan 17 10:46:53 2016
 */
/* Compiler settings for TangramCOMEx.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_ICOMHelper,0x7831ED8B,0xCDF1,0x4676,0x8D,0x26,0xE0,0xF6,0x07,0x32,0x62,0xD8);


MIDL_DEFINE_GUID(IID, LIBID_TangramCOMExLib,0xDADFC698,0x6649,0x4167,0xA5,0x01,0xEB,0x57,0x15,0x3A,0x20,0x1B);


MIDL_DEFINE_GUID(IID, DIID__ICOMHelperEvents,0x9ADF6F60,0xB663,0x4BC0,0x92,0x56,0x73,0xA7,0xDC,0x34,0xF1,0x48);


MIDL_DEFINE_GUID(CLSID, CLSID_COMHelper,0xE2379B15,0xCA40,0x4233,0xA4,0x4E,0x75,0x8D,0xC8,0x91,0xBE,0x1B);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



