

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


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


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __TangramCOMEx_h__
#define __TangramCOMEx_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ICOMHelper_FWD_DEFINED__
#define __ICOMHelper_FWD_DEFINED__
typedef interface ICOMHelper ICOMHelper;

#endif 	/* __ICOMHelper_FWD_DEFINED__ */


#ifndef ___ICOMHelperEvents_FWD_DEFINED__
#define ___ICOMHelperEvents_FWD_DEFINED__
typedef interface _ICOMHelperEvents _ICOMHelperEvents;

#endif 	/* ___ICOMHelperEvents_FWD_DEFINED__ */


#ifndef __COMHelper_FWD_DEFINED__
#define __COMHelper_FWD_DEFINED__

#ifdef __cplusplus
typedef class COMHelper COMHelper;
#else
typedef struct COMHelper COMHelper;
#endif /* __cplusplus */

#endif 	/* __COMHelper_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "shobjidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __ICOMHelper_INTERFACE_DEFINED__
#define __ICOMHelper_INTERFACE_DEFINED__

/* interface ICOMHelper */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICOMHelper;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("7831ED8B-CDF1-4676-8D26-E0F6073262D8")
    ICOMHelper : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE HookCOM( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE UnHookCOM( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE HookCreateInstance( 
            VARIANT_BOOL bHook) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE HookCoRegisterClassObject( 
            VARIANT_BOOL bHook) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICOMHelperVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICOMHelper * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICOMHelper * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICOMHelper * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICOMHelper * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICOMHelper * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICOMHelper * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICOMHelper * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *HookCOM )( 
            ICOMHelper * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *UnHookCOM )( 
            ICOMHelper * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *HookCreateInstance )( 
            ICOMHelper * This,
            VARIANT_BOOL bHook);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *HookCoRegisterClassObject )( 
            ICOMHelper * This,
            VARIANT_BOOL bHook);
        
        END_INTERFACE
    } ICOMHelperVtbl;

    interface ICOMHelper
    {
        CONST_VTBL struct ICOMHelperVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICOMHelper_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICOMHelper_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICOMHelper_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICOMHelper_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICOMHelper_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICOMHelper_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICOMHelper_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICOMHelper_HookCOM(This)	\
    ( (This)->lpVtbl -> HookCOM(This) ) 

#define ICOMHelper_UnHookCOM(This)	\
    ( (This)->lpVtbl -> UnHookCOM(This) ) 

#define ICOMHelper_HookCreateInstance(This,bHook)	\
    ( (This)->lpVtbl -> HookCreateInstance(This,bHook) ) 

#define ICOMHelper_HookCoRegisterClassObject(This,bHook)	\
    ( (This)->lpVtbl -> HookCoRegisterClassObject(This,bHook) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICOMHelper_INTERFACE_DEFINED__ */



#ifndef __TangramCOMExLib_LIBRARY_DEFINED__
#define __TangramCOMExLib_LIBRARY_DEFINED__

/* library TangramCOMExLib */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_TangramCOMExLib;

#ifndef ___ICOMHelperEvents_DISPINTERFACE_DEFINED__
#define ___ICOMHelperEvents_DISPINTERFACE_DEFINED__

/* dispinterface _ICOMHelperEvents */
/* [uuid] */ 


EXTERN_C const IID DIID__ICOMHelperEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("9ADF6F60-B663-4BC0-9256-73A7DC34F148")
    _ICOMHelperEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _ICOMHelperEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _ICOMHelperEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _ICOMHelperEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _ICOMHelperEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _ICOMHelperEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _ICOMHelperEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _ICOMHelperEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _ICOMHelperEvents * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _ICOMHelperEventsVtbl;

    interface _ICOMHelperEvents
    {
        CONST_VTBL struct _ICOMHelperEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _ICOMHelperEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _ICOMHelperEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _ICOMHelperEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _ICOMHelperEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _ICOMHelperEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _ICOMHelperEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _ICOMHelperEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___ICOMHelperEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_COMHelper;

#ifdef __cplusplus

class DECLSPEC_UUID("E2379B15-CA40-4233-A44E-758DC891BE1B")
COMHelper;
#endif
#endif /* __TangramCOMExLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


