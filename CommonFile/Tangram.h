

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Mon Jun 27 05:03:11 2016
 */
/* Compiler settings for ..\CommonFile\Tangram.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
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

#ifndef __Tangram_h__
#define __Tangram_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IRestNotify_FWD_DEFINED__
#define __IRestNotify_FWD_DEFINED__
typedef interface IRestNotify IRestNotify;

#endif 	/* __IRestNotify_FWD_DEFINED__ */


#ifndef __ITaskNotify_FWD_DEFINED__
#define __ITaskNotify_FWD_DEFINED__
typedef interface ITaskNotify ITaskNotify;

#endif 	/* __ITaskNotify_FWD_DEFINED__ */


#ifndef __ITask_FWD_DEFINED__
#define __ITask_FWD_DEFINED__
typedef interface ITask ITask;

#endif 	/* __ITask_FWD_DEFINED__ */


#ifndef __ITangramTreeNode_FWD_DEFINED__
#define __ITangramTreeNode_FWD_DEFINED__
typedef interface ITangramTreeNode ITangramTreeNode;

#endif 	/* __ITangramTreeNode_FWD_DEFINED__ */


#ifndef __ITangramTreeViewCallBack_FWD_DEFINED__
#define __ITangramTreeViewCallBack_FWD_DEFINED__
typedef interface ITangramTreeViewCallBack ITangramTreeViewCallBack;

#endif 	/* __ITangramTreeViewCallBack_FWD_DEFINED__ */


#ifndef __ITangramTreeView_FWD_DEFINED__
#define __ITangramTreeView_FWD_DEFINED__
typedef interface ITangramTreeView ITangramTreeView;

#endif 	/* __ITangramTreeView_FWD_DEFINED__ */


#ifndef __ITangramApp_FWD_DEFINED__
#define __ITangramApp_FWD_DEFINED__
typedef interface ITangramApp ITangramApp;

#endif 	/* __ITangramApp_FWD_DEFINED__ */


#ifndef __IEventProxy_FWD_DEFINED__
#define __IEventProxy_FWD_DEFINED__
typedef interface IEventProxy IEventProxy;

#endif 	/* __IEventProxy_FWD_DEFINED__ */


#ifndef __IWndNode_FWD_DEFINED__
#define __IWndNode_FWD_DEFINED__
typedef interface IWndNode IWndNode;

#endif 	/* __IWndNode_FWD_DEFINED__ */


#ifndef __IAppExtender_FWD_DEFINED__
#define __IAppExtender_FWD_DEFINED__
typedef interface IAppExtender IAppExtender;

#endif 	/* __IAppExtender_FWD_DEFINED__ */


#ifndef __IWndContainer_FWD_DEFINED__
#define __IWndContainer_FWD_DEFINED__
typedef interface IWndContainer IWndContainer;

#endif 	/* __IWndContainer_FWD_DEFINED__ */


#ifndef __ICreator_FWD_DEFINED__
#define __ICreator_FWD_DEFINED__
typedef interface ICreator ICreator;

#endif 	/* __ICreator_FWD_DEFINED__ */


#ifndef __ITangramEditor_FWD_DEFINED__
#define __ITangramEditor_FWD_DEFINED__
typedef interface ITangramEditor ITangramEditor;

#endif 	/* __ITangramEditor_FWD_DEFINED__ */


#ifndef __IJavaProxy_FWD_DEFINED__
#define __IJavaProxy_FWD_DEFINED__
typedef interface IJavaProxy IJavaProxy;

#endif 	/* __IJavaProxy_FWD_DEFINED__ */


#ifndef __ITangramAddin_FWD_DEFINED__
#define __ITangramAddin_FWD_DEFINED__
typedef interface ITangramAddin ITangramAddin;

#endif 	/* __ITangramAddin_FWD_DEFINED__ */


#ifndef __ITangram_FWD_DEFINED__
#define __ITangram_FWD_DEFINED__
typedef interface ITangram ITangram;

#endif 	/* __ITangram_FWD_DEFINED__ */


#ifndef __IWndFrame_FWD_DEFINED__
#define __IWndFrame_FWD_DEFINED__
typedef interface IWndFrame IWndFrame;

#endif 	/* __IWndFrame_FWD_DEFINED__ */


#ifndef __IWndNodeCollection_FWD_DEFINED__
#define __IWndNodeCollection_FWD_DEFINED__
typedef interface IWndNodeCollection IWndNodeCollection;

#endif 	/* __IWndNodeCollection_FWD_DEFINED__ */


#ifndef __IWndPage_FWD_DEFINED__
#define __IWndPage_FWD_DEFINED__
typedef interface IWndPage IWndPage;

#endif 	/* __IWndPage_FWD_DEFINED__ */


#ifndef __IEclipseTopWnd_FWD_DEFINED__
#define __IEclipseTopWnd_FWD_DEFINED__
typedef interface IEclipseTopWnd IEclipseTopWnd;

#endif 	/* __IEclipseTopWnd_FWD_DEFINED__ */


#ifndef __IOfficeObject_FWD_DEFINED__
#define __IOfficeObject_FWD_DEFINED__
typedef interface IOfficeObject IOfficeObject;

#endif 	/* __IOfficeObject_FWD_DEFINED__ */


#ifndef __ITaskObj_FWD_DEFINED__
#define __ITaskObj_FWD_DEFINED__
typedef interface ITaskObj ITaskObj;

#endif 	/* __ITaskObj_FWD_DEFINED__ */


#ifndef __IWebTaskObj_FWD_DEFINED__
#define __IWebTaskObj_FWD_DEFINED__
typedef interface IWebTaskObj IWebTaskObj;

#endif 	/* __IWebTaskObj_FWD_DEFINED__ */


#ifndef __IRestObject_FWD_DEFINED__
#define __IRestObject_FWD_DEFINED__
typedef interface IRestObject IRestObject;

#endif 	/* __IRestObject_FWD_DEFINED__ */


#ifndef __ITangramBot_FWD_DEFINED__
#define __ITangramBot_FWD_DEFINED__
typedef interface ITangramBot ITangramBot;

#endif 	/* __ITangramBot_FWD_DEFINED__ */


#ifndef __ICollaborationProxy_FWD_DEFINED__
#define __ICollaborationProxy_FWD_DEFINED__
typedef interface ICollaborationProxy ICollaborationProxy;

#endif 	/* __ICollaborationProxy_FWD_DEFINED__ */


#ifndef __ITangramCtrl_FWD_DEFINED__
#define __ITangramCtrl_FWD_DEFINED__
typedef interface ITangramCtrl ITangramCtrl;

#endif 	/* __ITangramCtrl_FWD_DEFINED__ */


#ifndef ___ITangramEvents_FWD_DEFINED__
#define ___ITangramEvents_FWD_DEFINED__
typedef interface _ITangramEvents _ITangramEvents;

#endif 	/* ___ITangramEvents_FWD_DEFINED__ */


#ifndef __Tangram_FWD_DEFINED__
#define __Tangram_FWD_DEFINED__

#ifdef __cplusplus
typedef class Tangram Tangram;
#else
typedef struct Tangram Tangram;
#endif /* __cplusplus */

#endif 	/* __Tangram_FWD_DEFINED__ */


#ifndef ___IWndNodeEvents_FWD_DEFINED__
#define ___IWndNodeEvents_FWD_DEFINED__
typedef interface _IWndNodeEvents _IWndNodeEvents;

#endif 	/* ___IWndNodeEvents_FWD_DEFINED__ */


#ifndef ___IWndPage_FWD_DEFINED__
#define ___IWndPage_FWD_DEFINED__
typedef interface _IWndPage _IWndPage;

#endif 	/* ___IWndPage_FWD_DEFINED__ */


#ifndef ___IEventProxy_FWD_DEFINED__
#define ___IEventProxy_FWD_DEFINED__
typedef interface _IEventProxy _IEventProxy;

#endif 	/* ___IEventProxy_FWD_DEFINED__ */


#ifndef ___IOfficeObjectEvents_FWD_DEFINED__
#define ___IOfficeObjectEvents_FWD_DEFINED__
typedef interface _IOfficeObjectEvents _IOfficeObjectEvents;

#endif 	/* ___IOfficeObjectEvents_FWD_DEFINED__ */


#ifndef __OfficeObject_FWD_DEFINED__
#define __OfficeObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class OfficeObject OfficeObject;
#else
typedef struct OfficeObject OfficeObject;
#endif /* __cplusplus */

#endif 	/* __OfficeObject_FWD_DEFINED__ */


#ifndef ___ITangramBotEvents_FWD_DEFINED__
#define ___ITangramBotEvents_FWD_DEFINED__
typedef interface _ITangramBotEvents _ITangramBotEvents;

#endif 	/* ___ITangramBotEvents_FWD_DEFINED__ */


#ifndef __TangramBot_FWD_DEFINED__
#define __TangramBot_FWD_DEFINED__

#ifdef __cplusplus
typedef class TangramBot TangramBot;
#else
typedef struct TangramBot TangramBot;
#endif /* __cplusplus */

#endif 	/* __TangramBot_FWD_DEFINED__ */


#ifndef __CollaborationProxy_FWD_DEFINED__
#define __CollaborationProxy_FWD_DEFINED__

#ifdef __cplusplus
typedef class CollaborationProxy CollaborationProxy;
#else
typedef struct CollaborationProxy CollaborationProxy;
#endif /* __cplusplus */

#endif 	/* __CollaborationProxy_FWD_DEFINED__ */


#ifndef __TangramCtrl_FWD_DEFINED__
#define __TangramCtrl_FWD_DEFINED__

#ifdef __cplusplus
typedef class TangramCtrl TangramCtrl;
#else
typedef struct TangramCtrl TangramCtrl;
#endif /* __cplusplus */

#endif 	/* __TangramCtrl_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "shobjidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_Tangram_0000_0000 */
/* [local] */ 












typedef /* [helpstring] */ 
enum WindowEventType
    {
        TangramClick	= 0,
        TangramDoubleClick	= 0x1,
        TangramEnter	= 0x2,
        TangramLeave	= 0x3,
        TangramEnabledChanged	= 0x4,
        TangramLostFocus	= 0x5,
        TangramGotFocus	= 0x6,
        TangramKeyUp	= 0x7,
        TangramKeyDown	= 0x8,
        TangramKeyPress	= 0x9,
        TangramMouseClick	= 0xa,
        TangramMouseDoubleClick	= 0xb,
        TangramMouseDown	= 0xc,
        TangramMouseEnter	= 0xd,
        TangramMouseHover	= 0xe,
        TangramMouseLeave	= 0xf,
        TangramMouseMove	= 0x10,
        TangramMouseUp	= 0x11,
        TangramMouseWheel	= 0x12,
        TangramTextChanged	= 0x13,
        TangramVisibleChanged	= 0x14,
        TangramClientSizeChanged	= 0x15,
        TangramSizeChanged	= 0x16,
        TangramParentChanged	= 0x17,
        TangramResize	= 0x18
    } 	WindowEventType;

typedef 
enum TaskNodeType
    {
        TANGRAM	= 0,
        UCMA	= 0x1,
        TaskNoWait	= 0x2,
        TaskNeedWait	= 0x3,
        DefaultNode	= 0x4
    } 	TaskNodeType;

typedef /* [helpstring] */ 
enum WndNodeType
    {
        TNT_Blank	= 0x1,
        TNT_ActiveX	= 0x2,
        TNT_Splitter	= 0x4,
        TNT_Tabbed	= 0x8,
        TNT_CLR_Control	= 0x10,
        TNT_CLR_Form	= 0x20,
        TNT_CLR_Window	= 0x40,
        TNT_View	= 0x80,
        TNT_TreeView	= 0xb0
    } 	WndNodeType;



extern RPC_IF_HANDLE __MIDL_itf_Tangram_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_Tangram_0000_0000_v0_0_s_ifspec;

#ifndef __IRestNotify_INTERFACE_DEFINED__
#define __IRestNotify_INTERFACE_DEFINED__

/* interface IRestNotify */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IRestNotify;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-06011982F7CD")
    IRestNotify : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Notify( 
            BSTR bstrInfo) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IRestNotifyVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IRestNotify * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IRestNotify * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IRestNotify * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IRestNotify * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IRestNotify * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IRestNotify * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IRestNotify * This,
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
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Notify )( 
            IRestNotify * This,
            BSTR bstrInfo);
        
        END_INTERFACE
    } IRestNotifyVtbl;

    interface IRestNotify
    {
        CONST_VTBL struct IRestNotifyVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IRestNotify_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IRestNotify_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IRestNotify_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IRestNotify_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IRestNotify_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IRestNotify_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IRestNotify_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IRestNotify_Notify(This,bstrInfo)	\
    ( (This)->lpVtbl -> Notify(This,bstrInfo) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IRestNotify_INTERFACE_DEFINED__ */


#ifndef __ITaskNotify_INTERFACE_DEFINED__
#define __ITaskNotify_INTERFACE_DEFINED__

/* interface ITaskNotify */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITaskNotify;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-0601198222C0")
    ITaskNotify : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Notify( 
            BSTR bstrInfo) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE NotifyEx( 
            VARIANT varNotify) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ITaskNotifyVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITaskNotify * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITaskNotify * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITaskNotify * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITaskNotify * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITaskNotify * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITaskNotify * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITaskNotify * This,
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
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Notify )( 
            ITaskNotify * This,
            BSTR bstrInfo);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *NotifyEx )( 
            ITaskNotify * This,
            VARIANT varNotify);
        
        END_INTERFACE
    } ITaskNotifyVtbl;

    interface ITaskNotify
    {
        CONST_VTBL struct ITaskNotifyVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITaskNotify_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITaskNotify_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITaskNotify_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITaskNotify_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITaskNotify_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITaskNotify_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITaskNotify_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITaskNotify_Notify(This,bstrInfo)	\
    ( (This)->lpVtbl -> Notify(This,bstrInfo) ) 

#define ITaskNotify_NotifyEx(This,varNotify)	\
    ( (This)->lpVtbl -> NotifyEx(This,varNotify) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITaskNotify_INTERFACE_DEFINED__ */


#ifndef __ITask_INTERFACE_DEFINED__
#define __ITask_INTERFACE_DEFINED__

/* interface ITask */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITask;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-0601198222C1")
    ITask : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Execute( 
            BSTR bstrInfo,
            ITaskNotify *pTangramTaskNotify,
            IDispatch *pStateDisp,
            ITask *pPrevTangramTask) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ITaskVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITask * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITask * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITask * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITask * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITask * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITask * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITask * This,
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
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Execute )( 
            ITask * This,
            BSTR bstrInfo,
            ITaskNotify *pTangramTaskNotify,
            IDispatch *pStateDisp,
            ITask *pPrevTangramTask);
        
        END_INTERFACE
    } ITaskVtbl;

    interface ITask
    {
        CONST_VTBL struct ITaskVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITask_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITask_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITask_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITask_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITask_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITask_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITask_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITask_Execute(This,bstrInfo,pTangramTaskNotify,pStateDisp,pPrevTangramTask)	\
    ( (This)->lpVtbl -> Execute(This,bstrInfo,pTangramTaskNotify,pStateDisp,pPrevTangramTask) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITask_INTERFACE_DEFINED__ */


#ifndef __ITangramTreeNode_INTERFACE_DEFINED__
#define __ITangramTreeNode_INTERFACE_DEFINED__

/* interface ITangramTreeNode */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITangramTreeNode;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-06011982C951")
    ITangramTreeNode : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct ITangramTreeNodeVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITangramTreeNode * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITangramTreeNode * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITangramTreeNode * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITangramTreeNode * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITangramTreeNode * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITangramTreeNode * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITangramTreeNode * This,
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
    } ITangramTreeNodeVtbl;

    interface ITangramTreeNode
    {
        CONST_VTBL struct ITangramTreeNodeVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITangramTreeNode_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITangramTreeNode_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITangramTreeNode_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITangramTreeNode_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITangramTreeNode_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITangramTreeNode_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITangramTreeNode_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITangramTreeNode_INTERFACE_DEFINED__ */


#ifndef __ITangramTreeViewCallBack_INTERFACE_DEFINED__
#define __ITangramTreeViewCallBack_INTERFACE_DEFINED__

/* interface ITangramTreeViewCallBack */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITangramTreeViewCallBack;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-0601198283A6")
    ITangramTreeViewCallBack : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_WndNode( 
            IWndNode *newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Pages( 
            long *retVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OnInitTreeView( 
            ITangramTreeView *pTangramTreeView,
            BSTR bstrXml) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OnClick( 
            BSTR bstrXml,
            BSTR bstrXmlData) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OnInit( 
            BSTR bstrXml) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OnNewPage( 
            int nNewPage) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TangramAction( 
            BSTR bstrXml,
            /* [retval][out] */ BSTR *bstrRetXml) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ITangramTreeViewCallBackVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITangramTreeViewCallBack * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITangramTreeViewCallBack * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITangramTreeViewCallBack * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITangramTreeViewCallBack * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITangramTreeViewCallBack * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITangramTreeViewCallBack * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITangramTreeViewCallBack * This,
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
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_WndNode )( 
            ITangramTreeViewCallBack * This,
            IWndNode *newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Pages )( 
            ITangramTreeViewCallBack * This,
            long *retVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *OnInitTreeView )( 
            ITangramTreeViewCallBack * This,
            ITangramTreeView *pTangramTreeView,
            BSTR bstrXml);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *OnClick )( 
            ITangramTreeViewCallBack * This,
            BSTR bstrXml,
            BSTR bstrXmlData);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *OnInit )( 
            ITangramTreeViewCallBack * This,
            BSTR bstrXml);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *OnNewPage )( 
            ITangramTreeViewCallBack * This,
            int nNewPage);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *TangramAction )( 
            ITangramTreeViewCallBack * This,
            BSTR bstrXml,
            /* [retval][out] */ BSTR *bstrRetXml);
        
        END_INTERFACE
    } ITangramTreeViewCallBackVtbl;

    interface ITangramTreeViewCallBack
    {
        CONST_VTBL struct ITangramTreeViewCallBackVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITangramTreeViewCallBack_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITangramTreeViewCallBack_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITangramTreeViewCallBack_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITangramTreeViewCallBack_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITangramTreeViewCallBack_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITangramTreeViewCallBack_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITangramTreeViewCallBack_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITangramTreeViewCallBack_put_WndNode(This,newVal)	\
    ( (This)->lpVtbl -> put_WndNode(This,newVal) ) 

#define ITangramTreeViewCallBack_get_Pages(This,retVal)	\
    ( (This)->lpVtbl -> get_Pages(This,retVal) ) 

#define ITangramTreeViewCallBack_OnInitTreeView(This,pTangramTreeView,bstrXml)	\
    ( (This)->lpVtbl -> OnInitTreeView(This,pTangramTreeView,bstrXml) ) 

#define ITangramTreeViewCallBack_OnClick(This,bstrXml,bstrXmlData)	\
    ( (This)->lpVtbl -> OnClick(This,bstrXml,bstrXmlData) ) 

#define ITangramTreeViewCallBack_OnInit(This,bstrXml)	\
    ( (This)->lpVtbl -> OnInit(This,bstrXml) ) 

#define ITangramTreeViewCallBack_OnNewPage(This,nNewPage)	\
    ( (This)->lpVtbl -> OnNewPage(This,nNewPage) ) 

#define ITangramTreeViewCallBack_TangramAction(This,bstrXml,bstrRetXml)	\
    ( (This)->lpVtbl -> TangramAction(This,bstrXml,bstrRetXml) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITangramTreeViewCallBack_INTERFACE_DEFINED__ */


#ifndef __ITangramTreeView_INTERFACE_DEFINED__
#define __ITangramTreeView_INTERFACE_DEFINED__

/* interface ITangramTreeView */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITangramTreeView;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-0601198283A5")
    ITangramTreeView : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_TangramTreeViewCallBack( 
            BSTR bstrKey,
            /* [in] */ ITangramTreeViewCallBack *newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_FirstRoot( 
            /* [retval][out] */ long *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddTreeNode( 
            long hItem,
            BSTR bstrXml) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE InsertNode( 
            BSTR bstrXml,
            /* [retval][out] */ long *hItem) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ITangramTreeViewVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITangramTreeView * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITangramTreeView * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITangramTreeView * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITangramTreeView * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITangramTreeView * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITangramTreeView * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITangramTreeView * This,
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
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_TangramTreeViewCallBack )( 
            ITangramTreeView * This,
            BSTR bstrKey,
            /* [in] */ ITangramTreeViewCallBack *newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_FirstRoot )( 
            ITangramTreeView * This,
            /* [retval][out] */ long *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *AddTreeNode )( 
            ITangramTreeView * This,
            long hItem,
            BSTR bstrXml);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *InsertNode )( 
            ITangramTreeView * This,
            BSTR bstrXml,
            /* [retval][out] */ long *hItem);
        
        END_INTERFACE
    } ITangramTreeViewVtbl;

    interface ITangramTreeView
    {
        CONST_VTBL struct ITangramTreeViewVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITangramTreeView_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITangramTreeView_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITangramTreeView_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITangramTreeView_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITangramTreeView_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITangramTreeView_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITangramTreeView_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITangramTreeView_put_TangramTreeViewCallBack(This,bstrKey,newVal)	\
    ( (This)->lpVtbl -> put_TangramTreeViewCallBack(This,bstrKey,newVal) ) 

#define ITangramTreeView_get_FirstRoot(This,pVal)	\
    ( (This)->lpVtbl -> get_FirstRoot(This,pVal) ) 

#define ITangramTreeView_AddTreeNode(This,hItem,bstrXml)	\
    ( (This)->lpVtbl -> AddTreeNode(This,hItem,bstrXml) ) 

#define ITangramTreeView_InsertNode(This,bstrXml,hItem)	\
    ( (This)->lpVtbl -> InsertNode(This,bstrXml,hItem) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITangramTreeView_INTERFACE_DEFINED__ */


#ifndef __ITangramApp_INTERFACE_DEFINED__
#define __ITangramApp_INTERFACE_DEFINED__

/* interface ITangramApp */
/* [unique][helpstring][nonextensible][hidden][dual][uuid][object] */ 


EXTERN_C const IID IID_ITangramApp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119820000")
    ITangramApp : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Tangram( 
            /* [retval][out] */ ITangram **ppTangramCore) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ITangramAppVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITangramApp * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITangramApp * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITangramApp * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITangramApp * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITangramApp * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITangramApp * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITangramApp * This,
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
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Tangram )( 
            ITangramApp * This,
            /* [retval][out] */ ITangram **ppTangramCore);
        
        END_INTERFACE
    } ITangramAppVtbl;

    interface ITangramApp
    {
        CONST_VTBL struct ITangramAppVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITangramApp_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITangramApp_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITangramApp_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITangramApp_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITangramApp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITangramApp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITangramApp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITangramApp_get_Tangram(This,ppTangramCore)	\
    ( (This)->lpVtbl -> get_Tangram(This,ppTangramCore) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITangramApp_INTERFACE_DEFINED__ */


#ifndef __IEventProxy_INTERFACE_DEFINED__
#define __IEventProxy_INTERFACE_DEFINED__

/* interface IEventProxy */
/* [unique][helpstring][nonextensible][hidden][dual][uuid][object] */ 


EXTERN_C const IID IID_IEventProxy;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119820005")
    IEventProxy : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IEventProxyVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IEventProxy * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IEventProxy * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IEventProxy * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IEventProxy * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IEventProxy * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IEventProxy * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IEventProxy * This,
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
    } IEventProxyVtbl;

    interface IEventProxy
    {
        CONST_VTBL struct IEventProxyVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IEventProxy_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IEventProxy_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IEventProxy_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IEventProxy_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IEventProxy_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IEventProxy_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IEventProxy_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IEventProxy_INTERFACE_DEFINED__ */


#ifndef __IWndNode_INTERFACE_DEFINED__
#define __IWndNode_INTERFACE_DEFINED__

/* interface IWndNode */
/* [object][unique][helpstring][uuid] */ 


EXTERN_C const IID IID_IWndNode;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119820004")
    IWndNode : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ChildNodes( 
            /* [retval][out] */ IWndNodeCollection **ppNodeColletion) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Rows( 
            /* [retval][out] */ long *nRows) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Cols( 
            /* [retval][out] */ long *nCols) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Row( 
            /* [retval][out] */ long *nRow) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Col( 
            /* [retval][out] */ long *nCol) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NodeType( 
            /* [retval][out] */ WndNodeType *nType) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ParentNode( 
            /* [retval][out] */ IWndNode **ppNode) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HTMLWindow( 
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_WndPage( 
            /* [retval][out] */ IWndPage **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RootNode( 
            /* [retval][out] */ IWndNode **ppNode) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_XObject( 
            /* [retval][out] */ VARIANT *pVar) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_AxPlugIn( 
            /* [in] */ BSTR bstrPlugInName,
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Caption( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Caption( 
            /* [in] */ BSTR bstrCaption) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ BSTR bstrName) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Objects( 
            /* [in] */ long nType,
            /* [retval][out] */ IWndNodeCollection **ppNodeColletion) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Handle( 
            /* [retval][out] */ long *hWnd) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Attribute( 
            /* [in] */ BSTR bstrKey,
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Attribute( 
            /* [in] */ BSTR bstrKey,
            /* [in] */ BSTR bstrVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Tag( 
            /* [retval][out] */ VARIANT *pVar) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Tag( 
            /* [in] */ VARIANT vVar) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_OuterXml( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Key( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HtmlDocument( 
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_NameAtWindowPage( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Width( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Height( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Frame( 
            /* [retval][out] */ IWndFrame **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_XML( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Extender( 
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Extender( 
            /* [in] */ IDispatch *newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_DocXml( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_rgbMiddle( 
            /* [retval][out] */ OLE_COLOR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_rgbMiddle( 
            /* [in] */ OLE_COLOR newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_rgbRightBottom( 
            /* [retval][out] */ OLE_COLOR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_rgbRightBottom( 
            /* [in] */ OLE_COLOR newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_rgbLeftTop( 
            /* [retval][out] */ OLE_COLOR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_rgbLeftTop( 
            /* [in] */ OLE_COLOR newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Hmin( 
            /* [retval][out] */ int *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Hmin( 
            /* [in] */ int newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Hmax( 
            /* [retval][out] */ int *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Hmax( 
            /* [in] */ int newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Vmin( 
            /* [retval][out] */ int *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Vmin( 
            /* [in] */ int newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Vmax( 
            /* [retval][out] */ int *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Vmax( 
            /* [in] */ int newVal) = 0;
        
        virtual /* [hidden][id] */ HRESULT STDMETHODCALLTYPE ActiveTabPage( 
            IWndNode *pNode) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetNodes( 
            /* [in] */ BSTR bstrName,
            /* [out] */ IWndNode **ppNode,
            /* [out] */ IWndNodeCollection **ppNodes,
            /* [retval][out] */ long *pCount) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetNode( 
            /* [in] */ long nRow,
            /* [in] */ long nCol,
            /* [retval][out] */ IWndNode **ppWndmNode) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetCtrlByName( 
            BSTR bstrName,
            VARIANT_BOOL bFindInChild,
            /* [retval][out] */ IDispatch **ppCtrlDisp) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Refresh( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Extend( 
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LoadXML( 
            int nType,
            BSTR bstrXML) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ExtendEx( 
            int nRow,
            int nCol,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE UpdateDesignerData( 
            BSTR bstrXML) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IWndNodeVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IWndNode * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IWndNode * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IWndNode * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IWndNode * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IWndNode * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IWndNode * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IWndNode * This,
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
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ChildNodes )( 
            IWndNode * This,
            /* [retval][out] */ IWndNodeCollection **ppNodeColletion);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Rows )( 
            IWndNode * This,
            /* [retval][out] */ long *nRows);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Cols )( 
            IWndNode * This,
            /* [retval][out] */ long *nCols);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Row )( 
            IWndNode * This,
            /* [retval][out] */ long *nRow);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Col )( 
            IWndNode * This,
            /* [retval][out] */ long *nCol);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_NodeType )( 
            IWndNode * This,
            /* [retval][out] */ WndNodeType *nType);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ParentNode )( 
            IWndNode * This,
            /* [retval][out] */ IWndNode **ppNode);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HTMLWindow )( 
            IWndNode * This,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_WndPage )( 
            IWndNode * This,
            /* [retval][out] */ IWndPage **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_RootNode )( 
            IWndNode * This,
            /* [retval][out] */ IWndNode **ppNode);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_XObject )( 
            IWndNode * This,
            /* [retval][out] */ VARIANT *pVar);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_AxPlugIn )( 
            IWndNode * This,
            /* [in] */ BSTR bstrPlugInName,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Caption )( 
            IWndNode * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Caption )( 
            IWndNode * This,
            /* [in] */ BSTR bstrCaption);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Name )( 
            IWndNode * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Name )( 
            IWndNode * This,
            /* [in] */ BSTR bstrName);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Objects )( 
            IWndNode * This,
            /* [in] */ long nType,
            /* [retval][out] */ IWndNodeCollection **ppNodeColletion);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Handle )( 
            IWndNode * This,
            /* [retval][out] */ long *hWnd);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Attribute )( 
            IWndNode * This,
            /* [in] */ BSTR bstrKey,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Attribute )( 
            IWndNode * This,
            /* [in] */ BSTR bstrKey,
            /* [in] */ BSTR bstrVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Tag )( 
            IWndNode * This,
            /* [retval][out] */ VARIANT *pVar);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Tag )( 
            IWndNode * This,
            /* [in] */ VARIANT vVar);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OuterXml )( 
            IWndNode * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Key )( 
            IWndNode * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HtmlDocument )( 
            IWndNode * This,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_NameAtWindowPage )( 
            IWndNode * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Width )( 
            IWndNode * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Height )( 
            IWndNode * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Frame )( 
            IWndNode * This,
            /* [retval][out] */ IWndFrame **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_XML )( 
            IWndNode * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Extender )( 
            IWndNode * This,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Extender )( 
            IWndNode * This,
            /* [in] */ IDispatch *newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_DocXml )( 
            IWndNode * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_rgbMiddle )( 
            IWndNode * This,
            /* [retval][out] */ OLE_COLOR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_rgbMiddle )( 
            IWndNode * This,
            /* [in] */ OLE_COLOR newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_rgbRightBottom )( 
            IWndNode * This,
            /* [retval][out] */ OLE_COLOR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_rgbRightBottom )( 
            IWndNode * This,
            /* [in] */ OLE_COLOR newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_rgbLeftTop )( 
            IWndNode * This,
            /* [retval][out] */ OLE_COLOR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_rgbLeftTop )( 
            IWndNode * This,
            /* [in] */ OLE_COLOR newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Hmin )( 
            IWndNode * This,
            /* [retval][out] */ int *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Hmin )( 
            IWndNode * This,
            /* [in] */ int newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Hmax )( 
            IWndNode * This,
            /* [retval][out] */ int *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Hmax )( 
            IWndNode * This,
            /* [in] */ int newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Vmin )( 
            IWndNode * This,
            /* [retval][out] */ int *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Vmin )( 
            IWndNode * This,
            /* [in] */ int newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Vmax )( 
            IWndNode * This,
            /* [retval][out] */ int *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Vmax )( 
            IWndNode * This,
            /* [in] */ int newVal);
        
        /* [hidden][id] */ HRESULT ( STDMETHODCALLTYPE *ActiveTabPage )( 
            IWndNode * This,
            IWndNode *pNode);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetNodes )( 
            IWndNode * This,
            /* [in] */ BSTR bstrName,
            /* [out] */ IWndNode **ppNode,
            /* [out] */ IWndNodeCollection **ppNodes,
            /* [retval][out] */ long *pCount);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetNode )( 
            IWndNode * This,
            /* [in] */ long nRow,
            /* [in] */ long nCol,
            /* [retval][out] */ IWndNode **ppWndmNode);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetCtrlByName )( 
            IWndNode * This,
            BSTR bstrName,
            VARIANT_BOOL bFindInChild,
            /* [retval][out] */ IDispatch **ppCtrlDisp);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Refresh )( 
            IWndNode * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Extend )( 
            IWndNode * This,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *LoadXML )( 
            IWndNode * This,
            int nType,
            BSTR bstrXML);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ExtendEx )( 
            IWndNode * This,
            int nRow,
            int nCol,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *UpdateDesignerData )( 
            IWndNode * This,
            BSTR bstrXML);
        
        END_INTERFACE
    } IWndNodeVtbl;

    interface IWndNode
    {
        CONST_VTBL struct IWndNodeVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWndNode_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IWndNode_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IWndNode_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IWndNode_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IWndNode_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IWndNode_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IWndNode_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IWndNode_get_ChildNodes(This,ppNodeColletion)	\
    ( (This)->lpVtbl -> get_ChildNodes(This,ppNodeColletion) ) 

#define IWndNode_get_Rows(This,nRows)	\
    ( (This)->lpVtbl -> get_Rows(This,nRows) ) 

#define IWndNode_get_Cols(This,nCols)	\
    ( (This)->lpVtbl -> get_Cols(This,nCols) ) 

#define IWndNode_get_Row(This,nRow)	\
    ( (This)->lpVtbl -> get_Row(This,nRow) ) 

#define IWndNode_get_Col(This,nCol)	\
    ( (This)->lpVtbl -> get_Col(This,nCol) ) 

#define IWndNode_get_NodeType(This,nType)	\
    ( (This)->lpVtbl -> get_NodeType(This,nType) ) 

#define IWndNode_get_ParentNode(This,ppNode)	\
    ( (This)->lpVtbl -> get_ParentNode(This,ppNode) ) 

#define IWndNode_get_HTMLWindow(This,pVal)	\
    ( (This)->lpVtbl -> get_HTMLWindow(This,pVal) ) 

#define IWndNode_get_WndPage(This,pVal)	\
    ( (This)->lpVtbl -> get_WndPage(This,pVal) ) 

#define IWndNode_get_RootNode(This,ppNode)	\
    ( (This)->lpVtbl -> get_RootNode(This,ppNode) ) 

#define IWndNode_get_XObject(This,pVar)	\
    ( (This)->lpVtbl -> get_XObject(This,pVar) ) 

#define IWndNode_get_AxPlugIn(This,bstrPlugInName,pVal)	\
    ( (This)->lpVtbl -> get_AxPlugIn(This,bstrPlugInName,pVal) ) 

#define IWndNode_get_Caption(This,pVal)	\
    ( (This)->lpVtbl -> get_Caption(This,pVal) ) 

#define IWndNode_put_Caption(This,bstrCaption)	\
    ( (This)->lpVtbl -> put_Caption(This,bstrCaption) ) 

#define IWndNode_get_Name(This,pVal)	\
    ( (This)->lpVtbl -> get_Name(This,pVal) ) 

#define IWndNode_put_Name(This,bstrName)	\
    ( (This)->lpVtbl -> put_Name(This,bstrName) ) 

#define IWndNode_get_Objects(This,nType,ppNodeColletion)	\
    ( (This)->lpVtbl -> get_Objects(This,nType,ppNodeColletion) ) 

#define IWndNode_get_Handle(This,hWnd)	\
    ( (This)->lpVtbl -> get_Handle(This,hWnd) ) 

#define IWndNode_get_Attribute(This,bstrKey,pVal)	\
    ( (This)->lpVtbl -> get_Attribute(This,bstrKey,pVal) ) 

#define IWndNode_put_Attribute(This,bstrKey,bstrVal)	\
    ( (This)->lpVtbl -> put_Attribute(This,bstrKey,bstrVal) ) 

#define IWndNode_get_Tag(This,pVar)	\
    ( (This)->lpVtbl -> get_Tag(This,pVar) ) 

#define IWndNode_put_Tag(This,vVar)	\
    ( (This)->lpVtbl -> put_Tag(This,vVar) ) 

#define IWndNode_get_OuterXml(This,pVal)	\
    ( (This)->lpVtbl -> get_OuterXml(This,pVal) ) 

#define IWndNode_get_Key(This,pVal)	\
    ( (This)->lpVtbl -> get_Key(This,pVal) ) 

#define IWndNode_get_HtmlDocument(This,pVal)	\
    ( (This)->lpVtbl -> get_HtmlDocument(This,pVal) ) 

#define IWndNode_get_NameAtWindowPage(This,pVal)	\
    ( (This)->lpVtbl -> get_NameAtWindowPage(This,pVal) ) 

#define IWndNode_get_Width(This,pVal)	\
    ( (This)->lpVtbl -> get_Width(This,pVal) ) 

#define IWndNode_get_Height(This,pVal)	\
    ( (This)->lpVtbl -> get_Height(This,pVal) ) 

#define IWndNode_get_Frame(This,pVal)	\
    ( (This)->lpVtbl -> get_Frame(This,pVal) ) 

#define IWndNode_get_XML(This,pVal)	\
    ( (This)->lpVtbl -> get_XML(This,pVal) ) 

#define IWndNode_get_Extender(This,pVal)	\
    ( (This)->lpVtbl -> get_Extender(This,pVal) ) 

#define IWndNode_put_Extender(This,newVal)	\
    ( (This)->lpVtbl -> put_Extender(This,newVal) ) 

#define IWndNode_get_DocXml(This,pVal)	\
    ( (This)->lpVtbl -> get_DocXml(This,pVal) ) 

#define IWndNode_get_rgbMiddle(This,pVal)	\
    ( (This)->lpVtbl -> get_rgbMiddle(This,pVal) ) 

#define IWndNode_put_rgbMiddle(This,newVal)	\
    ( (This)->lpVtbl -> put_rgbMiddle(This,newVal) ) 

#define IWndNode_get_rgbRightBottom(This,pVal)	\
    ( (This)->lpVtbl -> get_rgbRightBottom(This,pVal) ) 

#define IWndNode_put_rgbRightBottom(This,newVal)	\
    ( (This)->lpVtbl -> put_rgbRightBottom(This,newVal) ) 

#define IWndNode_get_rgbLeftTop(This,pVal)	\
    ( (This)->lpVtbl -> get_rgbLeftTop(This,pVal) ) 

#define IWndNode_put_rgbLeftTop(This,newVal)	\
    ( (This)->lpVtbl -> put_rgbLeftTop(This,newVal) ) 

#define IWndNode_get_Hmin(This,pVal)	\
    ( (This)->lpVtbl -> get_Hmin(This,pVal) ) 

#define IWndNode_put_Hmin(This,newVal)	\
    ( (This)->lpVtbl -> put_Hmin(This,newVal) ) 

#define IWndNode_get_Hmax(This,pVal)	\
    ( (This)->lpVtbl -> get_Hmax(This,pVal) ) 

#define IWndNode_put_Hmax(This,newVal)	\
    ( (This)->lpVtbl -> put_Hmax(This,newVal) ) 

#define IWndNode_get_Vmin(This,pVal)	\
    ( (This)->lpVtbl -> get_Vmin(This,pVal) ) 

#define IWndNode_put_Vmin(This,newVal)	\
    ( (This)->lpVtbl -> put_Vmin(This,newVal) ) 

#define IWndNode_get_Vmax(This,pVal)	\
    ( (This)->lpVtbl -> get_Vmax(This,pVal) ) 

#define IWndNode_put_Vmax(This,newVal)	\
    ( (This)->lpVtbl -> put_Vmax(This,newVal) ) 

#define IWndNode_ActiveTabPage(This,pNode)	\
    ( (This)->lpVtbl -> ActiveTabPage(This,pNode) ) 

#define IWndNode_GetNodes(This,bstrName,ppNode,ppNodes,pCount)	\
    ( (This)->lpVtbl -> GetNodes(This,bstrName,ppNode,ppNodes,pCount) ) 

#define IWndNode_GetNode(This,nRow,nCol,ppWndmNode)	\
    ( (This)->lpVtbl -> GetNode(This,nRow,nCol,ppWndmNode) ) 

#define IWndNode_GetCtrlByName(This,bstrName,bFindInChild,ppCtrlDisp)	\
    ( (This)->lpVtbl -> GetCtrlByName(This,bstrName,bFindInChild,ppCtrlDisp) ) 

#define IWndNode_Refresh(This)	\
    ( (This)->lpVtbl -> Refresh(This) ) 

#define IWndNode_Extend(This,bstrKey,bstrXml,ppRetNode)	\
    ( (This)->lpVtbl -> Extend(This,bstrKey,bstrXml,ppRetNode) ) 

#define IWndNode_LoadXML(This,nType,bstrXML)	\
    ( (This)->lpVtbl -> LoadXML(This,nType,bstrXML) ) 

#define IWndNode_ExtendEx(This,nRow,nCol,bstrKey,bstrXml,ppRetNode)	\
    ( (This)->lpVtbl -> ExtendEx(This,nRow,nCol,bstrKey,bstrXml,ppRetNode) ) 

#define IWndNode_UpdateDesignerData(This,bstrXML)	\
    ( (This)->lpVtbl -> UpdateDesignerData(This,bstrXML) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IWndNode_INTERFACE_DEFINED__ */


#ifndef __IAppExtender_INTERFACE_DEFINED__
#define __IAppExtender_INTERFACE_DEFINED__

/* interface IAppExtender */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IAppExtender;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119822007")
    IAppExtender : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ProcessNotify( 
            /* [in] */ BSTR bstrXmlNotify) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IAppExtenderVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IAppExtender * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IAppExtender * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IAppExtender * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IAppExtender * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IAppExtender * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IAppExtender * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IAppExtender * This,
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
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ProcessNotify )( 
            IAppExtender * This,
            /* [in] */ BSTR bstrXmlNotify);
        
        END_INTERFACE
    } IAppExtenderVtbl;

    interface IAppExtender
    {
        CONST_VTBL struct IAppExtenderVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IAppExtender_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IAppExtender_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IAppExtender_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IAppExtender_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IAppExtender_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IAppExtender_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IAppExtender_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IAppExtender_ProcessNotify(This,bstrXmlNotify)	\
    ( (This)->lpVtbl -> ProcessNotify(This,bstrXmlNotify) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IAppExtender_INTERFACE_DEFINED__ */


#ifndef __IWndContainer_INTERFACE_DEFINED__
#define __IWndContainer_INTERFACE_DEFINED__

/* interface IWndContainer */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IWndContainer;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119820007")
    IWndContainer : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Save( void) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IWndContainerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IWndContainer * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IWndContainer * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IWndContainer * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IWndContainer * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IWndContainer * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IWndContainer * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IWndContainer * This,
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
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Save )( 
            IWndContainer * This);
        
        END_INTERFACE
    } IWndContainerVtbl;

    interface IWndContainer
    {
        CONST_VTBL struct IWndContainerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWndContainer_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IWndContainer_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IWndContainer_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IWndContainer_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IWndContainer_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IWndContainer_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IWndContainer_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IWndContainer_Save(This)	\
    ( (This)->lpVtbl -> Save(This) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IWndContainer_INTERFACE_DEFINED__ */


#ifndef __ICreator_INTERFACE_DEFINED__
#define __ICreator_INTERFACE_DEFINED__

/* interface ICreator */
/* [uuid][object] */ 


EXTERN_C const IID IID_ICreator;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119820006")
    ICreator : public IDispatch
    {
    public:
        virtual /* [hidden][helpstring][id] */ HRESULT STDMETHODCALLTYPE Create( 
            /* [in] */ long hParentWnd,
            /* [in] */ IWndNode *pNode,
            /* [out][in] */ long *hWnd) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Names( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Tags( 
            BSTR bstrObjName,
            /* [retval][out] */ BSTR *pVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICreatorVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICreator * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICreator * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICreator * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICreator * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICreator * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICreator * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICreator * This,
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
        
        /* [hidden][helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Create )( 
            ICreator * This,
            /* [in] */ long hParentWnd,
            /* [in] */ IWndNode *pNode,
            /* [out][in] */ long *hWnd);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Names )( 
            ICreator * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Tags )( 
            ICreator * This,
            BSTR bstrObjName,
            /* [retval][out] */ BSTR *pVal);
        
        END_INTERFACE
    } ICreatorVtbl;

    interface ICreator
    {
        CONST_VTBL struct ICreatorVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICreator_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICreator_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICreator_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICreator_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICreator_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICreator_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICreator_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICreator_Create(This,hParentWnd,pNode,hWnd)	\
    ( (This)->lpVtbl -> Create(This,hParentWnd,pNode,hWnd) ) 

#define ICreator_get_Names(This,pVal)	\
    ( (This)->lpVtbl -> get_Names(This,pVal) ) 

#define ICreator_get_Tags(This,bstrObjName,pVal)	\
    ( (This)->lpVtbl -> get_Tags(This,bstrObjName,pVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICreator_INTERFACE_DEFINED__ */


#ifndef __ITangramEditor_INTERFACE_DEFINED__
#define __ITangramEditor_INTERFACE_DEFINED__

/* interface ITangramEditor */
/* [unique][helpstring][nonextensible][oleautomation][uuid][object] */ 


EXTERN_C const IID IID_ITangramEditor;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119826688")
    ITangramEditor : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct ITangramEditorVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITangramEditor * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITangramEditor * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITangramEditor * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITangramEditor * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITangramEditor * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITangramEditor * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITangramEditor * This,
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
    } ITangramEditorVtbl;

    interface ITangramEditor
    {
        CONST_VTBL struct ITangramEditorVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITangramEditor_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITangramEditor_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITangramEditor_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITangramEditor_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITangramEditor_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITangramEditor_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITangramEditor_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITangramEditor_INTERFACE_DEFINED__ */


#ifndef __IJavaProxy_INTERFACE_DEFINED__
#define __IJavaProxy_INTERFACE_DEFINED__

/* interface IJavaProxy */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IJavaProxy;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119827001")
    IJavaProxy : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE InitEclipse( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE InitComProxy( void) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IJavaProxyVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IJavaProxy * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IJavaProxy * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IJavaProxy * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IJavaProxy * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IJavaProxy * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IJavaProxy * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IJavaProxy * This,
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
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *InitEclipse )( 
            IJavaProxy * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *InitComProxy )( 
            IJavaProxy * This);
        
        END_INTERFACE
    } IJavaProxyVtbl;

    interface IJavaProxy
    {
        CONST_VTBL struct IJavaProxyVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IJavaProxy_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IJavaProxy_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IJavaProxy_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IJavaProxy_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IJavaProxy_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IJavaProxy_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IJavaProxy_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IJavaProxy_InitEclipse(This)	\
    ( (This)->lpVtbl -> InitEclipse(This) ) 

#define IJavaProxy_InitComProxy(This)	\
    ( (This)->lpVtbl -> InitComProxy(This) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IJavaProxy_INTERFACE_DEFINED__ */


#ifndef __ITangramAddin_INTERFACE_DEFINED__
#define __ITangramAddin_INTERFACE_DEFINED__

/* interface ITangramAddin */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITangramAddin;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119828001")
    ITangramAddin : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct ITangramAddinVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITangramAddin * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITangramAddin * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITangramAddin * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITangramAddin * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITangramAddin * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITangramAddin * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITangramAddin * This,
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
    } ITangramAddinVtbl;

    interface ITangramAddin
    {
        CONST_VTBL struct ITangramAddinVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITangramAddin_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITangramAddin_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITangramAddin_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITangramAddin_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITangramAddin_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITangramAddin_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITangramAddin_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITangramAddin_INTERFACE_DEFINED__ */


#ifndef __ITangram_INTERFACE_DEFINED__
#define __ITangram_INTERFACE_DEFINED__

/* interface ITangram */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITangram;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119820001")
    ITangram : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_AppKeyValue( 
            BSTR bstrKey,
            /* [retval][out] */ VARIANT *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_AppKeyValue( 
            BSTR bstrKey,
            /* [in] */ VARIANT newVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ExternalInfo( 
            /* [in] */ VARIANT newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_AppExtender( 
            BSTR bstrKey,
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_AppExtender( 
            BSTR bstrKey,
            /* [in] */ IDispatch *newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RootNodes( 
            /* [retval][out] */ IWndNodeCollection **ppNodeColletion) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Application( 
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CurrentActiveNode( 
            /* [retval][out] */ IWndNode **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CreatingNode( 
            /* [retval][out] */ IWndNode **pVal) = 0;
        
        virtual /* [hidden][id][propput] */ HRESULT STDMETHODCALLTYPE put_CurrentDesignNode( 
            /* [in] */ IWndNode *newVal) = 0;
        
        virtual /* [hidden][id][propget] */ HRESULT STDMETHODCALLTYPE get_DesignNode( 
            /* [retval][out] */ IWndNode **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_RemoteHelperHWND( 
            /* [retval][out] */ long *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_RemoteHelperHWND( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_JavaProxy( 
            /* [in] */ IJavaProxy *newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HostWnd( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_CollaborationProxy( 
            /* [in] */ ICollaborationProxy *newVal) = 0;
        
        virtual /* [hidden][id] */ HRESULT STDMETHODCALLTYPE GetNewLayoutNodeName( 
            BSTR strCnnID,
            IWndNode *pDesignNode,
            /* [retval][out] */ BSTR *bstrNewName) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateWndPage( 
            long hWnd,
            /* [retval][out] */ IWndPage **ppTangram) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateRestObj( 
            /* [retval][out] */ IRestObject **ppRestObj) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateTaskObj( 
            /* [retval][out] */ ITaskObj **ppTaskObj) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateWebTask( 
            BSTR bstrTaskName,
            BSTR bstrTaskURL,
            /* [retval][out] */ IWebTaskObj **pWebTaskObj) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CreateCLRObj( 
            BSTR bstrObjID,
            /* [retval][out] */ IDispatch **ppDisp) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetCLRControlString( 
            BSTR bstrAssemblyPath,
            /* [retval][out] */ BSTR *bstrCtrls) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetWndFrame( 
            long hHostWnd,
            /* [retval][out] */ IWndFrame **ppFrame) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetItemText( 
            IWndNode *pNode,
            long nCtrlID,
            LONG nMaxLengeh,
            /* [retval][out] */ BSTR *bstrRet) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SetItemText( 
            IWndNode *pNode,
            long nCtrlID,
            BSTR bstrText) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetCLRControl( 
            IDispatch *CtrlDisp,
            BSTR bstrNames,
            /* [retval][out] */ IDispatch **ppRetDisp) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE MessageBox( 
            long hWnd,
            BSTR bstrContext,
            BSTR bstrCaption,
            long nStyle,
            /* [retval][out] */ int *nRet) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Encode( 
            BSTR bstrSRC,
            VARIANT_BOOL bEncode,
            /* [retval][out] */ BSTR *bstrRet) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetHostFocus( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE UpdateWndNode( 
            IWndNode *pNode) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE NewGUID( 
            /* [retval][out] */ BSTR *retVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ActiveCLRMethod( 
            BSTR bstrObjID,
            BSTR bstrMethod,
            BSTR bstrParam,
            BSTR bstrData) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TangramGetObject( 
            IDispatch *SourceDisp,
            IDispatch **ResultDisp) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateOfficeObj( 
            BSTR bstrAppID,
            BSTR bstrXml) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateOfficeDocument( 
            BSTR bstrXml) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE DownLoadFile( 
            BSTR strFileURL,
            BSTR bstrTargetFile,
            BSTR bstrActionXml) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ExtendXml( 
            BSTR bstrXml,
            BSTR bstrKey,
            /* [retval][out] */ IDispatch **ppNode) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetNodeFromeHandle( 
            long hWnd,
            /* [retval][out] */ IWndNode **ppRetNode) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE InitVBAForm( 
            /* [in] */ IDispatch *newVal,
            /* [in] */ long nStyle,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppNode) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddDocXml( 
            IDispatch *pDocdisp,
            BSTR bstrXml,
            BSTR bstrKey,
            /* [retval][out] */ VARIANT_BOOL *bSuccess) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetDocXmlByKey( 
            IDispatch *pDocdisp,
            BSTR bstrKey,
            /* [retval][out] */ BSTR *pbstrRet) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AddVBAFormsScript( 
            IDispatch *OfficeObject,
            BSTR bstrKey,
            BSTR bstrXml) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ExportOfficeObjXml( 
            IDispatch *OfficeObject,
            /* [retval][out] */ BSTR *bstrXml) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetFrameFromVBAForm( 
            IDispatch *pForm,
            /* [retval][out] */ IWndFrame **ppFrame) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetActiveTopWndNode( 
            IDispatch *pForm,
            /* [retval][out] */ IWndNode **WndNode) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetObjectFromWnd( 
            LONG hWnd,
            /* [retval][out] */ IDispatch **ppObjFromWnd) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ReleaseTangram( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE InitJava( 
            int nIndex) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TangramCommand( 
            IDispatch *RibbonControl) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TangramGetImage( 
            BSTR strValue,
            /* [retval][out] */ IPictureDisp **ppdispImage) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TangramGetVisible( 
            IDispatch *RibbonControl,
            /* [retval][out] */ VARIANT *varVisible) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TangramOnLoad( 
            IDispatch *RibbonControl) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TangramGetItemCount( 
            IDispatch *RibbonControl,
            /* [retval][out] */ long *nCount) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TangramGetItemLabel( 
            IDispatch *RibbonControl,
            long nIndex,
            /* [retval][out] */ BSTR *bstrLabel) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TangramGetItemID( 
            IDispatch *RibbonControl,
            long nIndex,
            /* [retval][out] */ BSTR *bstrID) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ITangramVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITangram * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITangram * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITangram * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITangram * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITangram * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITangram * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITangram * This,
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
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_AppKeyValue )( 
            ITangram * This,
            BSTR bstrKey,
            /* [retval][out] */ VARIANT *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_AppKeyValue )( 
            ITangram * This,
            BSTR bstrKey,
            /* [in] */ VARIANT newVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ExternalInfo )( 
            ITangram * This,
            /* [in] */ VARIANT newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_AppExtender )( 
            ITangram * This,
            BSTR bstrKey,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_AppExtender )( 
            ITangram * This,
            BSTR bstrKey,
            /* [in] */ IDispatch *newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_RootNodes )( 
            ITangram * This,
            /* [retval][out] */ IWndNodeCollection **ppNodeColletion);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Application )( 
            ITangram * This,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_CurrentActiveNode )( 
            ITangram * This,
            /* [retval][out] */ IWndNode **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_CreatingNode )( 
            ITangram * This,
            /* [retval][out] */ IWndNode **pVal);
        
        /* [hidden][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_CurrentDesignNode )( 
            ITangram * This,
            /* [in] */ IWndNode *newVal);
        
        /* [hidden][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_DesignNode )( 
            ITangram * This,
            /* [retval][out] */ IWndNode **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_RemoteHelperHWND )( 
            ITangram * This,
            /* [retval][out] */ long *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_RemoteHelperHWND )( 
            ITangram * This,
            /* [in] */ long newVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_JavaProxy )( 
            ITangram * This,
            /* [in] */ IJavaProxy *newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HostWnd )( 
            ITangram * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_CollaborationProxy )( 
            ITangram * This,
            /* [in] */ ICollaborationProxy *newVal);
        
        /* [hidden][id] */ HRESULT ( STDMETHODCALLTYPE *GetNewLayoutNodeName )( 
            ITangram * This,
            BSTR strCnnID,
            IWndNode *pDesignNode,
            /* [retval][out] */ BSTR *bstrNewName);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CreateWndPage )( 
            ITangram * This,
            long hWnd,
            /* [retval][out] */ IWndPage **ppTangram);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CreateRestObj )( 
            ITangram * This,
            /* [retval][out] */ IRestObject **ppRestObj);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CreateTaskObj )( 
            ITangram * This,
            /* [retval][out] */ ITaskObj **ppTaskObj);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CreateWebTask )( 
            ITangram * This,
            BSTR bstrTaskName,
            BSTR bstrTaskURL,
            /* [retval][out] */ IWebTaskObj **pWebTaskObj);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CreateCLRObj )( 
            ITangram * This,
            BSTR bstrObjID,
            /* [retval][out] */ IDispatch **ppDisp);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetCLRControlString )( 
            ITangram * This,
            BSTR bstrAssemblyPath,
            /* [retval][out] */ BSTR *bstrCtrls);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetWndFrame )( 
            ITangram * This,
            long hHostWnd,
            /* [retval][out] */ IWndFrame **ppFrame);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetItemText )( 
            ITangram * This,
            IWndNode *pNode,
            long nCtrlID,
            LONG nMaxLengeh,
            /* [retval][out] */ BSTR *bstrRet);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *SetItemText )( 
            ITangram * This,
            IWndNode *pNode,
            long nCtrlID,
            BSTR bstrText);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetCLRControl )( 
            ITangram * This,
            IDispatch *CtrlDisp,
            BSTR bstrNames,
            /* [retval][out] */ IDispatch **ppRetDisp);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *MessageBox )( 
            ITangram * This,
            long hWnd,
            BSTR bstrContext,
            BSTR bstrCaption,
            long nStyle,
            /* [retval][out] */ int *nRet);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Encode )( 
            ITangram * This,
            BSTR bstrSRC,
            VARIANT_BOOL bEncode,
            /* [retval][out] */ BSTR *bstrRet);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SetHostFocus )( 
            ITangram * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *UpdateWndNode )( 
            ITangram * This,
            IWndNode *pNode);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *NewGUID )( 
            ITangram * This,
            /* [retval][out] */ BSTR *retVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ActiveCLRMethod )( 
            ITangram * This,
            BSTR bstrObjID,
            BSTR bstrMethod,
            BSTR bstrParam,
            BSTR bstrData);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *TangramGetObject )( 
            ITangram * This,
            IDispatch *SourceDisp,
            IDispatch **ResultDisp);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CreateOfficeObj )( 
            ITangram * This,
            BSTR bstrAppID,
            BSTR bstrXml);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CreateOfficeDocument )( 
            ITangram * This,
            BSTR bstrXml);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *DownLoadFile )( 
            ITangram * This,
            BSTR strFileURL,
            BSTR bstrTargetFile,
            BSTR bstrActionXml);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ExtendXml )( 
            ITangram * This,
            BSTR bstrXml,
            BSTR bstrKey,
            /* [retval][out] */ IDispatch **ppNode);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetNodeFromeHandle )( 
            ITangram * This,
            long hWnd,
            /* [retval][out] */ IWndNode **ppRetNode);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *InitVBAForm )( 
            ITangram * This,
            /* [in] */ IDispatch *newVal,
            /* [in] */ long nStyle,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppNode);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *AddDocXml )( 
            ITangram * This,
            IDispatch *pDocdisp,
            BSTR bstrXml,
            BSTR bstrKey,
            /* [retval][out] */ VARIANT_BOOL *bSuccess);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetDocXmlByKey )( 
            ITangram * This,
            IDispatch *pDocdisp,
            BSTR bstrKey,
            /* [retval][out] */ BSTR *pbstrRet);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AddVBAFormsScript )( 
            ITangram * This,
            IDispatch *OfficeObject,
            BSTR bstrKey,
            BSTR bstrXml);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ExportOfficeObjXml )( 
            ITangram * This,
            IDispatch *OfficeObject,
            /* [retval][out] */ BSTR *bstrXml);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetFrameFromVBAForm )( 
            ITangram * This,
            IDispatch *pForm,
            /* [retval][out] */ IWndFrame **ppFrame);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetActiveTopWndNode )( 
            ITangram * This,
            IDispatch *pForm,
            /* [retval][out] */ IWndNode **WndNode);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetObjectFromWnd )( 
            ITangram * This,
            LONG hWnd,
            /* [retval][out] */ IDispatch **ppObjFromWnd);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ReleaseTangram )( 
            ITangram * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *InitJava )( 
            ITangram * This,
            int nIndex);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *TangramCommand )( 
            ITangram * This,
            IDispatch *RibbonControl);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *TangramGetImage )( 
            ITangram * This,
            BSTR strValue,
            /* [retval][out] */ IPictureDisp **ppdispImage);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *TangramGetVisible )( 
            ITangram * This,
            IDispatch *RibbonControl,
            /* [retval][out] */ VARIANT *varVisible);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *TangramOnLoad )( 
            ITangram * This,
            IDispatch *RibbonControl);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *TangramGetItemCount )( 
            ITangram * This,
            IDispatch *RibbonControl,
            /* [retval][out] */ long *nCount);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *TangramGetItemLabel )( 
            ITangram * This,
            IDispatch *RibbonControl,
            long nIndex,
            /* [retval][out] */ BSTR *bstrLabel);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *TangramGetItemID )( 
            ITangram * This,
            IDispatch *RibbonControl,
            long nIndex,
            /* [retval][out] */ BSTR *bstrID);
        
        END_INTERFACE
    } ITangramVtbl;

    interface ITangram
    {
        CONST_VTBL struct ITangramVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITangram_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITangram_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITangram_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITangram_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITangram_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITangram_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITangram_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITangram_get_AppKeyValue(This,bstrKey,pVal)	\
    ( (This)->lpVtbl -> get_AppKeyValue(This,bstrKey,pVal) ) 

#define ITangram_put_AppKeyValue(This,bstrKey,newVal)	\
    ( (This)->lpVtbl -> put_AppKeyValue(This,bstrKey,newVal) ) 

#define ITangram_put_ExternalInfo(This,newVal)	\
    ( (This)->lpVtbl -> put_ExternalInfo(This,newVal) ) 

#define ITangram_get_AppExtender(This,bstrKey,pVal)	\
    ( (This)->lpVtbl -> get_AppExtender(This,bstrKey,pVal) ) 

#define ITangram_put_AppExtender(This,bstrKey,newVal)	\
    ( (This)->lpVtbl -> put_AppExtender(This,bstrKey,newVal) ) 

#define ITangram_get_RootNodes(This,ppNodeColletion)	\
    ( (This)->lpVtbl -> get_RootNodes(This,ppNodeColletion) ) 

#define ITangram_get_Application(This,pVal)	\
    ( (This)->lpVtbl -> get_Application(This,pVal) ) 

#define ITangram_get_CurrentActiveNode(This,pVal)	\
    ( (This)->lpVtbl -> get_CurrentActiveNode(This,pVal) ) 

#define ITangram_get_CreatingNode(This,pVal)	\
    ( (This)->lpVtbl -> get_CreatingNode(This,pVal) ) 

#define ITangram_put_CurrentDesignNode(This,newVal)	\
    ( (This)->lpVtbl -> put_CurrentDesignNode(This,newVal) ) 

#define ITangram_get_DesignNode(This,pVal)	\
    ( (This)->lpVtbl -> get_DesignNode(This,pVal) ) 

#define ITangram_get_RemoteHelperHWND(This,pVal)	\
    ( (This)->lpVtbl -> get_RemoteHelperHWND(This,pVal) ) 

#define ITangram_put_RemoteHelperHWND(This,newVal)	\
    ( (This)->lpVtbl -> put_RemoteHelperHWND(This,newVal) ) 

#define ITangram_put_JavaProxy(This,newVal)	\
    ( (This)->lpVtbl -> put_JavaProxy(This,newVal) ) 

#define ITangram_get_HostWnd(This,pVal)	\
    ( (This)->lpVtbl -> get_HostWnd(This,pVal) ) 

#define ITangram_put_CollaborationProxy(This,newVal)	\
    ( (This)->lpVtbl -> put_CollaborationProxy(This,newVal) ) 

#define ITangram_GetNewLayoutNodeName(This,strCnnID,pDesignNode,bstrNewName)	\
    ( (This)->lpVtbl -> GetNewLayoutNodeName(This,strCnnID,pDesignNode,bstrNewName) ) 

#define ITangram_CreateWndPage(This,hWnd,ppTangram)	\
    ( (This)->lpVtbl -> CreateWndPage(This,hWnd,ppTangram) ) 

#define ITangram_CreateRestObj(This,ppRestObj)	\
    ( (This)->lpVtbl -> CreateRestObj(This,ppRestObj) ) 

#define ITangram_CreateTaskObj(This,ppTaskObj)	\
    ( (This)->lpVtbl -> CreateTaskObj(This,ppTaskObj) ) 

#define ITangram_CreateWebTask(This,bstrTaskName,bstrTaskURL,pWebTaskObj)	\
    ( (This)->lpVtbl -> CreateWebTask(This,bstrTaskName,bstrTaskURL,pWebTaskObj) ) 

#define ITangram_CreateCLRObj(This,bstrObjID,ppDisp)	\
    ( (This)->lpVtbl -> CreateCLRObj(This,bstrObjID,ppDisp) ) 

#define ITangram_GetCLRControlString(This,bstrAssemblyPath,bstrCtrls)	\
    ( (This)->lpVtbl -> GetCLRControlString(This,bstrAssemblyPath,bstrCtrls) ) 

#define ITangram_GetWndFrame(This,hHostWnd,ppFrame)	\
    ( (This)->lpVtbl -> GetWndFrame(This,hHostWnd,ppFrame) ) 

#define ITangram_GetItemText(This,pNode,nCtrlID,nMaxLengeh,bstrRet)	\
    ( (This)->lpVtbl -> GetItemText(This,pNode,nCtrlID,nMaxLengeh,bstrRet) ) 

#define ITangram_SetItemText(This,pNode,nCtrlID,bstrText)	\
    ( (This)->lpVtbl -> SetItemText(This,pNode,nCtrlID,bstrText) ) 

#define ITangram_GetCLRControl(This,CtrlDisp,bstrNames,ppRetDisp)	\
    ( (This)->lpVtbl -> GetCLRControl(This,CtrlDisp,bstrNames,ppRetDisp) ) 

#define ITangram_MessageBox(This,hWnd,bstrContext,bstrCaption,nStyle,nRet)	\
    ( (This)->lpVtbl -> MessageBox(This,hWnd,bstrContext,bstrCaption,nStyle,nRet) ) 

#define ITangram_Encode(This,bstrSRC,bEncode,bstrRet)	\
    ( (This)->lpVtbl -> Encode(This,bstrSRC,bEncode,bstrRet) ) 

#define ITangram_SetHostFocus(This)	\
    ( (This)->lpVtbl -> SetHostFocus(This) ) 

#define ITangram_UpdateWndNode(This,pNode)	\
    ( (This)->lpVtbl -> UpdateWndNode(This,pNode) ) 

#define ITangram_NewGUID(This,retVal)	\
    ( (This)->lpVtbl -> NewGUID(This,retVal) ) 

#define ITangram_ActiveCLRMethod(This,bstrObjID,bstrMethod,bstrParam,bstrData)	\
    ( (This)->lpVtbl -> ActiveCLRMethod(This,bstrObjID,bstrMethod,bstrParam,bstrData) ) 

#define ITangram_TangramGetObject(This,SourceDisp,ResultDisp)	\
    ( (This)->lpVtbl -> TangramGetObject(This,SourceDisp,ResultDisp) ) 

#define ITangram_CreateOfficeObj(This,bstrAppID,bstrXml)	\
    ( (This)->lpVtbl -> CreateOfficeObj(This,bstrAppID,bstrXml) ) 

#define ITangram_CreateOfficeDocument(This,bstrXml)	\
    ( (This)->lpVtbl -> CreateOfficeDocument(This,bstrXml) ) 

#define ITangram_DownLoadFile(This,strFileURL,bstrTargetFile,bstrActionXml)	\
    ( (This)->lpVtbl -> DownLoadFile(This,strFileURL,bstrTargetFile,bstrActionXml) ) 

#define ITangram_ExtendXml(This,bstrXml,bstrKey,ppNode)	\
    ( (This)->lpVtbl -> ExtendXml(This,bstrXml,bstrKey,ppNode) ) 

#define ITangram_GetNodeFromeHandle(This,hWnd,ppRetNode)	\
    ( (This)->lpVtbl -> GetNodeFromeHandle(This,hWnd,ppRetNode) ) 

#define ITangram_InitVBAForm(This,newVal,nStyle,bstrXml,ppNode)	\
    ( (This)->lpVtbl -> InitVBAForm(This,newVal,nStyle,bstrXml,ppNode) ) 

#define ITangram_AddDocXml(This,pDocdisp,bstrXml,bstrKey,bSuccess)	\
    ( (This)->lpVtbl -> AddDocXml(This,pDocdisp,bstrXml,bstrKey,bSuccess) ) 

#define ITangram_GetDocXmlByKey(This,pDocdisp,bstrKey,pbstrRet)	\
    ( (This)->lpVtbl -> GetDocXmlByKey(This,pDocdisp,bstrKey,pbstrRet) ) 

#define ITangram_AddVBAFormsScript(This,OfficeObject,bstrKey,bstrXml)	\
    ( (This)->lpVtbl -> AddVBAFormsScript(This,OfficeObject,bstrKey,bstrXml) ) 

#define ITangram_ExportOfficeObjXml(This,OfficeObject,bstrXml)	\
    ( (This)->lpVtbl -> ExportOfficeObjXml(This,OfficeObject,bstrXml) ) 

#define ITangram_GetFrameFromVBAForm(This,pForm,ppFrame)	\
    ( (This)->lpVtbl -> GetFrameFromVBAForm(This,pForm,ppFrame) ) 

#define ITangram_GetActiveTopWndNode(This,pForm,WndNode)	\
    ( (This)->lpVtbl -> GetActiveTopWndNode(This,pForm,WndNode) ) 

#define ITangram_GetObjectFromWnd(This,hWnd,ppObjFromWnd)	\
    ( (This)->lpVtbl -> GetObjectFromWnd(This,hWnd,ppObjFromWnd) ) 

#define ITangram_ReleaseTangram(This)	\
    ( (This)->lpVtbl -> ReleaseTangram(This) ) 

#define ITangram_InitJava(This,nIndex)	\
    ( (This)->lpVtbl -> InitJava(This,nIndex) ) 

#define ITangram_TangramCommand(This,RibbonControl)	\
    ( (This)->lpVtbl -> TangramCommand(This,RibbonControl) ) 

#define ITangram_TangramGetImage(This,strValue,ppdispImage)	\
    ( (This)->lpVtbl -> TangramGetImage(This,strValue,ppdispImage) ) 

#define ITangram_TangramGetVisible(This,RibbonControl,varVisible)	\
    ( (This)->lpVtbl -> TangramGetVisible(This,RibbonControl,varVisible) ) 

#define ITangram_TangramOnLoad(This,RibbonControl)	\
    ( (This)->lpVtbl -> TangramOnLoad(This,RibbonControl) ) 

#define ITangram_TangramGetItemCount(This,RibbonControl,nCount)	\
    ( (This)->lpVtbl -> TangramGetItemCount(This,RibbonControl,nCount) ) 

#define ITangram_TangramGetItemLabel(This,RibbonControl,nIndex,bstrLabel)	\
    ( (This)->lpVtbl -> TangramGetItemLabel(This,RibbonControl,nIndex,bstrLabel) ) 

#define ITangram_TangramGetItemID(This,RibbonControl,nIndex,bstrID)	\
    ( (This)->lpVtbl -> TangramGetItemID(This,RibbonControl,nIndex,bstrID) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITangram_INTERFACE_DEFINED__ */


#ifndef __IWndFrame_INTERFACE_DEFINED__
#define __IWndFrame_INTERFACE_DEFINED__

/* interface IWndFrame */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IWndFrame;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119820003")
    IWndFrame : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RootNodes( 
            /* [retval][out] */ IWndNodeCollection **ppNodeColletion) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_HWND( 
            /* [retval][out] */ long *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_VisibleNode( 
            /* [retval][out] */ IWndNode **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CurrentNavigateKey( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_WndPage( 
            /* [retval][out] */ IWndPage **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_FrameData( 
            BSTR bstrKey,
            /* [retval][out] */ VARIANT *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_FrameData( 
            BSTR bstrKey,
            /* [in] */ VARIANT newVal) = 0;
        
        virtual /* [hidden][id][propget] */ HRESULT STDMETHODCALLTYPE get_DesignerState( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [hidden][id][propput] */ HRESULT STDMETHODCALLTYPE put_DesignerState( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Detach( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Attach( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ModifyHost( 
            long hHostWnd) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Extend( 
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetXml( 
            BSTR bstrRootName,
            /* [retval][out] */ BSTR *bstrRet) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IWndFrameVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IWndFrame * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IWndFrame * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IWndFrame * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IWndFrame * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IWndFrame * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IWndFrame * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IWndFrame * This,
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
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_RootNodes )( 
            IWndFrame * This,
            /* [retval][out] */ IWndNodeCollection **ppNodeColletion);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HWND )( 
            IWndFrame * This,
            /* [retval][out] */ long *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_VisibleNode )( 
            IWndFrame * This,
            /* [retval][out] */ IWndNode **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_CurrentNavigateKey )( 
            IWndFrame * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_WndPage )( 
            IWndFrame * This,
            /* [retval][out] */ IWndPage **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_FrameData )( 
            IWndFrame * This,
            BSTR bstrKey,
            /* [retval][out] */ VARIANT *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_FrameData )( 
            IWndFrame * This,
            BSTR bstrKey,
            /* [in] */ VARIANT newVal);
        
        /* [hidden][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_DesignerState )( 
            IWndFrame * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        /* [hidden][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_DesignerState )( 
            IWndFrame * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Detach )( 
            IWndFrame * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Attach )( 
            IWndFrame * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ModifyHost )( 
            IWndFrame * This,
            long hHostWnd);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Extend )( 
            IWndFrame * This,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetXml )( 
            IWndFrame * This,
            BSTR bstrRootName,
            /* [retval][out] */ BSTR *bstrRet);
        
        END_INTERFACE
    } IWndFrameVtbl;

    interface IWndFrame
    {
        CONST_VTBL struct IWndFrameVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWndFrame_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IWndFrame_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IWndFrame_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IWndFrame_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IWndFrame_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IWndFrame_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IWndFrame_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IWndFrame_get_RootNodes(This,ppNodeColletion)	\
    ( (This)->lpVtbl -> get_RootNodes(This,ppNodeColletion) ) 

#define IWndFrame_get_HWND(This,pVal)	\
    ( (This)->lpVtbl -> get_HWND(This,pVal) ) 

#define IWndFrame_get_VisibleNode(This,pVal)	\
    ( (This)->lpVtbl -> get_VisibleNode(This,pVal) ) 

#define IWndFrame_get_CurrentNavigateKey(This,pVal)	\
    ( (This)->lpVtbl -> get_CurrentNavigateKey(This,pVal) ) 

#define IWndFrame_get_WndPage(This,pVal)	\
    ( (This)->lpVtbl -> get_WndPage(This,pVal) ) 

#define IWndFrame_get_FrameData(This,bstrKey,pVal)	\
    ( (This)->lpVtbl -> get_FrameData(This,bstrKey,pVal) ) 

#define IWndFrame_put_FrameData(This,bstrKey,newVal)	\
    ( (This)->lpVtbl -> put_FrameData(This,bstrKey,newVal) ) 

#define IWndFrame_get_DesignerState(This,pVal)	\
    ( (This)->lpVtbl -> get_DesignerState(This,pVal) ) 

#define IWndFrame_put_DesignerState(This,newVal)	\
    ( (This)->lpVtbl -> put_DesignerState(This,newVal) ) 

#define IWndFrame_Detach(This)	\
    ( (This)->lpVtbl -> Detach(This) ) 

#define IWndFrame_Attach(This)	\
    ( (This)->lpVtbl -> Attach(This) ) 

#define IWndFrame_ModifyHost(This,hHostWnd)	\
    ( (This)->lpVtbl -> ModifyHost(This,hHostWnd) ) 

#define IWndFrame_Extend(This,bstrKey,bstrXml,ppRetNode)	\
    ( (This)->lpVtbl -> Extend(This,bstrKey,bstrXml,ppRetNode) ) 

#define IWndFrame_GetXml(This,bstrRootName,bstrRet)	\
    ( (This)->lpVtbl -> GetXml(This,bstrRootName,bstrRet) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IWndFrame_INTERFACE_DEFINED__ */


#ifndef __IWndNodeCollection_INTERFACE_DEFINED__
#define __IWndNodeCollection_INTERFACE_DEFINED__

/* interface IWndNodeCollection */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IWndNodeCollection;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119820017")
    IWndNodeCollection : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NodeCount( 
            /* [retval][out] */ long *pCount) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ long iIndex,
            /* [retval][out] */ IWndNode **ppTopWindow) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **ppVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IWndNodeCollectionVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IWndNodeCollection * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IWndNodeCollection * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IWndNodeCollection * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IWndNodeCollection * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IWndNodeCollection * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IWndNodeCollection * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IWndNodeCollection * This,
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
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_NodeCount )( 
            IWndNodeCollection * This,
            /* [retval][out] */ long *pCount);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Item )( 
            IWndNodeCollection * This,
            /* [in] */ long iIndex,
            /* [retval][out] */ IWndNode **ppTopWindow);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get__NewEnum )( 
            IWndNodeCollection * This,
            /* [retval][out] */ IUnknown **ppVal);
        
        END_INTERFACE
    } IWndNodeCollectionVtbl;

    interface IWndNodeCollection
    {
        CONST_VTBL struct IWndNodeCollectionVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWndNodeCollection_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IWndNodeCollection_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IWndNodeCollection_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IWndNodeCollection_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IWndNodeCollection_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IWndNodeCollection_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IWndNodeCollection_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IWndNodeCollection_get_NodeCount(This,pCount)	\
    ( (This)->lpVtbl -> get_NodeCount(This,pCount) ) 

#define IWndNodeCollection_get_Item(This,iIndex,ppTopWindow)	\
    ( (This)->lpVtbl -> get_Item(This,iIndex,ppTopWindow) ) 

#define IWndNodeCollection_get__NewEnum(This,ppVal)	\
    ( (This)->lpVtbl -> get__NewEnum(This,ppVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IWndNodeCollection_INTERFACE_DEFINED__ */


#ifndef __IWndPage_INTERFACE_DEFINED__
#define __IWndPage_INTERFACE_DEFINED__

/* interface IWndPage */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IWndPage;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119820002")
    IWndPage : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Frame( 
            /* [in] */ VARIANT vIndex,
            /* [retval][out] */ IWndFrame **ppFrame) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **ppVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long *pCount) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_URL( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_URL( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableSinkCLRCtrlCreatedEvent( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableSinkCLRCtrlCreatedEvent( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_xtml( 
            BSTR strKey,
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_xtml( 
            BSTR strKey,
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Extender( 
            BSTR bstrExtenderName,
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Extender( 
            BSTR bstrExtenderName,
            /* [in] */ IDispatch *newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_FrameName( 
            long hHwnd,
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_WndFrame( 
            BSTR bstrFrameName,
            /* [retval][out] */ IWndFrame **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Nodes( 
            BSTR bstrNodeName,
            /* [retval][out] */ IWndNode **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_XObjects( 
            BSTR bstrName,
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HtmlDocument( 
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Width( 
            /* [retval][out] */ long *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Width( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Height( 
            /* [retval][out] */ long *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Height( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_NodeNames( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HTMLWindow( 
            BSTR NodeName,
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IWndPage **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_External( 
            /* [retval][out] */ IDispatch **ppVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_External( 
            /* [in] */ IDispatch *newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Handle( 
            /* [retval][out] */ long *pVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateFrame( 
            VARIANT ParentObj,
            VARIANT HostWnd,
            BSTR bstrFrameName,
            /* [retval][out] */ IWndFrame **pRetFrame) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AddObject( 
            IDispatch *pHtmlDoc,
            IDispatch *pObject,
            BSTR bstrObjName,
            VARIANT_BOOL bSinkEvent,
            /* [retval][out] */ VARIANT_BOOL *bResult) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AddObjToHtml( 
            BSTR strObjName,
            VARIANT_BOOL bConnectEvent,
            IDispatch *pObjDisp) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AddNodesToPage( 
            IDispatch *pHtmlDoc,
            BSTR bstrNodeNames,
            VARIANT_BOOL bAdd,
            /* [retval][out] */ VARIANT_BOOL *bSuccess) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AttachObjEvent( 
            IDispatch *HTMLWindow,
            IDispatch *SourceObj,
            BSTR bstrEventName,
            BSTR AliasName) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AttachNodeSubCtrlEvent( 
            IDispatch *HtmlWndDisp,
            VARIANT Node,
            VARIANT Ctrl,
            BSTR EventName,
            BSTR AliasName) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AttachEvent( 
            BSTR bstrNames,
            IDispatch *pHtmlWnd,
            /* [retval][out] */ VARIANT_BOOL *bResult) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AttachNodeEvent( 
            BSTR bstrNames,
            IDispatch *pWndDisp) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Extend( 
            IDispatch *Parent,
            BSTR CtrlName,
            BSTR FrameName,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ExtendCtrl( 
            VARIANT Ctrl,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetCtrlByName( 
            IDispatch *pCtrl,
            BSTR bstrName,
            VARIANT_BOOL bFindInChild,
            /* [retval][out] */ IDispatch **ppCtrlDisp) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetCtrlInNode( 
            BSTR NodeName,
            BSTR CtrlName,
            /* [retval][out] */ IDispatch **ppCtrl) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetHtmlDocument( 
            IDispatch *HtmlWindow,
            /* [retval][out] */ IDispatch **ppHtmlDoc) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetFrameFromCtrl( 
            IDispatch *ctrl,
            /* [retval][out] */ IWndFrame **ppFrame) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetWndNode( 
            BSTR bstrFrameName,
            BSTR bstrNodeName,
            /* [retval][out] */ IWndNode **pRetNode) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetCtrlValueByName( 
            IDispatch *pCtrl,
            BSTR bstrName,
            VARIANT_BOOL bFindInChild,
            /* [retval][out] */ BSTR *bstrVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IWndPageVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IWndPage * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IWndPage * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IWndPage * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IWndPage * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IWndPage * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IWndPage * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IWndPage * This,
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
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Frame )( 
            IWndPage * This,
            /* [in] */ VARIANT vIndex,
            /* [retval][out] */ IWndFrame **ppFrame);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get__NewEnum )( 
            IWndPage * This,
            /* [retval][out] */ IUnknown **ppVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Count )( 
            IWndPage * This,
            /* [retval][out] */ long *pCount);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_URL )( 
            IWndPage * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_URL )( 
            IWndPage * This,
            /* [in] */ BSTR newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_EnableSinkCLRCtrlCreatedEvent )( 
            IWndPage * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_EnableSinkCLRCtrlCreatedEvent )( 
            IWndPage * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_xtml )( 
            IWndPage * This,
            BSTR strKey,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_xtml )( 
            IWndPage * This,
            BSTR strKey,
            /* [in] */ BSTR newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Extender )( 
            IWndPage * This,
            BSTR bstrExtenderName,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Extender )( 
            IWndPage * This,
            BSTR bstrExtenderName,
            /* [in] */ IDispatch *newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_FrameName )( 
            IWndPage * This,
            long hHwnd,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_WndFrame )( 
            IWndPage * This,
            BSTR bstrFrameName,
            /* [retval][out] */ IWndFrame **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Nodes )( 
            IWndPage * This,
            BSTR bstrNodeName,
            /* [retval][out] */ IWndNode **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_XObjects )( 
            IWndPage * This,
            BSTR bstrName,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HtmlDocument )( 
            IWndPage * This,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Width )( 
            IWndPage * This,
            /* [retval][out] */ long *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Width )( 
            IWndPage * This,
            /* [in] */ long newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Height )( 
            IWndPage * This,
            /* [retval][out] */ long *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Height )( 
            IWndPage * This,
            /* [in] */ long newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_NodeNames )( 
            IWndPage * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HTMLWindow )( 
            IWndPage * This,
            BSTR NodeName,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Parent )( 
            IWndPage * This,
            /* [retval][out] */ IWndPage **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_External )( 
            IWndPage * This,
            /* [retval][out] */ IDispatch **ppVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_External )( 
            IWndPage * This,
            /* [in] */ IDispatch *newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Handle )( 
            IWndPage * This,
            /* [retval][out] */ long *pVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CreateFrame )( 
            IWndPage * This,
            VARIANT ParentObj,
            VARIANT HostWnd,
            BSTR bstrFrameName,
            /* [retval][out] */ IWndFrame **pRetFrame);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AddObject )( 
            IWndPage * This,
            IDispatch *pHtmlDoc,
            IDispatch *pObject,
            BSTR bstrObjName,
            VARIANT_BOOL bSinkEvent,
            /* [retval][out] */ VARIANT_BOOL *bResult);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AddObjToHtml )( 
            IWndPage * This,
            BSTR strObjName,
            VARIANT_BOOL bConnectEvent,
            IDispatch *pObjDisp);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AddNodesToPage )( 
            IWndPage * This,
            IDispatch *pHtmlDoc,
            BSTR bstrNodeNames,
            VARIANT_BOOL bAdd,
            /* [retval][out] */ VARIANT_BOOL *bSuccess);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AttachObjEvent )( 
            IWndPage * This,
            IDispatch *HTMLWindow,
            IDispatch *SourceObj,
            BSTR bstrEventName,
            BSTR AliasName);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AttachNodeSubCtrlEvent )( 
            IWndPage * This,
            IDispatch *HtmlWndDisp,
            VARIANT Node,
            VARIANT Ctrl,
            BSTR EventName,
            BSTR AliasName);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AttachEvent )( 
            IWndPage * This,
            BSTR bstrNames,
            IDispatch *pHtmlWnd,
            /* [retval][out] */ VARIANT_BOOL *bResult);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AttachNodeEvent )( 
            IWndPage * This,
            BSTR bstrNames,
            IDispatch *pWndDisp);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Extend )( 
            IWndPage * This,
            IDispatch *Parent,
            BSTR CtrlName,
            BSTR FrameName,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ExtendCtrl )( 
            IWndPage * This,
            VARIANT Ctrl,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetCtrlByName )( 
            IWndPage * This,
            IDispatch *pCtrl,
            BSTR bstrName,
            VARIANT_BOOL bFindInChild,
            /* [retval][out] */ IDispatch **ppCtrlDisp);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetCtrlInNode )( 
            IWndPage * This,
            BSTR NodeName,
            BSTR CtrlName,
            /* [retval][out] */ IDispatch **ppCtrl);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetHtmlDocument )( 
            IWndPage * This,
            IDispatch *HtmlWindow,
            /* [retval][out] */ IDispatch **ppHtmlDoc);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetFrameFromCtrl )( 
            IWndPage * This,
            IDispatch *ctrl,
            /* [retval][out] */ IWndFrame **ppFrame);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetWndNode )( 
            IWndPage * This,
            BSTR bstrFrameName,
            BSTR bstrNodeName,
            /* [retval][out] */ IWndNode **pRetNode);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetCtrlValueByName )( 
            IWndPage * This,
            IDispatch *pCtrl,
            BSTR bstrName,
            VARIANT_BOOL bFindInChild,
            /* [retval][out] */ BSTR *bstrVal);
        
        END_INTERFACE
    } IWndPageVtbl;

    interface IWndPage
    {
        CONST_VTBL struct IWndPageVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWndPage_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IWndPage_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IWndPage_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IWndPage_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IWndPage_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IWndPage_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IWndPage_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IWndPage_get_Frame(This,vIndex,ppFrame)	\
    ( (This)->lpVtbl -> get_Frame(This,vIndex,ppFrame) ) 

#define IWndPage_get__NewEnum(This,ppVal)	\
    ( (This)->lpVtbl -> get__NewEnum(This,ppVal) ) 

#define IWndPage_get_Count(This,pCount)	\
    ( (This)->lpVtbl -> get_Count(This,pCount) ) 

#define IWndPage_get_URL(This,pVal)	\
    ( (This)->lpVtbl -> get_URL(This,pVal) ) 

#define IWndPage_put_URL(This,newVal)	\
    ( (This)->lpVtbl -> put_URL(This,newVal) ) 

#define IWndPage_get_EnableSinkCLRCtrlCreatedEvent(This,pVal)	\
    ( (This)->lpVtbl -> get_EnableSinkCLRCtrlCreatedEvent(This,pVal) ) 

#define IWndPage_put_EnableSinkCLRCtrlCreatedEvent(This,newVal)	\
    ( (This)->lpVtbl -> put_EnableSinkCLRCtrlCreatedEvent(This,newVal) ) 

#define IWndPage_get_xtml(This,strKey,pVal)	\
    ( (This)->lpVtbl -> get_xtml(This,strKey,pVal) ) 

#define IWndPage_put_xtml(This,strKey,newVal)	\
    ( (This)->lpVtbl -> put_xtml(This,strKey,newVal) ) 

#define IWndPage_get_Extender(This,bstrExtenderName,pVal)	\
    ( (This)->lpVtbl -> get_Extender(This,bstrExtenderName,pVal) ) 

#define IWndPage_put_Extender(This,bstrExtenderName,newVal)	\
    ( (This)->lpVtbl -> put_Extender(This,bstrExtenderName,newVal) ) 

#define IWndPage_get_FrameName(This,hHwnd,pVal)	\
    ( (This)->lpVtbl -> get_FrameName(This,hHwnd,pVal) ) 

#define IWndPage_get_WndFrame(This,bstrFrameName,pVal)	\
    ( (This)->lpVtbl -> get_WndFrame(This,bstrFrameName,pVal) ) 

#define IWndPage_get_Nodes(This,bstrNodeName,pVal)	\
    ( (This)->lpVtbl -> get_Nodes(This,bstrNodeName,pVal) ) 

#define IWndPage_get_XObjects(This,bstrName,pVal)	\
    ( (This)->lpVtbl -> get_XObjects(This,bstrName,pVal) ) 

#define IWndPage_get_HtmlDocument(This,pVal)	\
    ( (This)->lpVtbl -> get_HtmlDocument(This,pVal) ) 

#define IWndPage_get_Width(This,pVal)	\
    ( (This)->lpVtbl -> get_Width(This,pVal) ) 

#define IWndPage_put_Width(This,newVal)	\
    ( (This)->lpVtbl -> put_Width(This,newVal) ) 

#define IWndPage_get_Height(This,pVal)	\
    ( (This)->lpVtbl -> get_Height(This,pVal) ) 

#define IWndPage_put_Height(This,newVal)	\
    ( (This)->lpVtbl -> put_Height(This,newVal) ) 

#define IWndPage_get_NodeNames(This,pVal)	\
    ( (This)->lpVtbl -> get_NodeNames(This,pVal) ) 

#define IWndPage_get_HTMLWindow(This,NodeName,pVal)	\
    ( (This)->lpVtbl -> get_HTMLWindow(This,NodeName,pVal) ) 

#define IWndPage_get_Parent(This,pVal)	\
    ( (This)->lpVtbl -> get_Parent(This,pVal) ) 

#define IWndPage_get_External(This,ppVal)	\
    ( (This)->lpVtbl -> get_External(This,ppVal) ) 

#define IWndPage_put_External(This,newVal)	\
    ( (This)->lpVtbl -> put_External(This,newVal) ) 

#define IWndPage_get_Handle(This,pVal)	\
    ( (This)->lpVtbl -> get_Handle(This,pVal) ) 

#define IWndPage_CreateFrame(This,ParentObj,HostWnd,bstrFrameName,pRetFrame)	\
    ( (This)->lpVtbl -> CreateFrame(This,ParentObj,HostWnd,bstrFrameName,pRetFrame) ) 

#define IWndPage_AddObject(This,pHtmlDoc,pObject,bstrObjName,bSinkEvent,bResult)	\
    ( (This)->lpVtbl -> AddObject(This,pHtmlDoc,pObject,bstrObjName,bSinkEvent,bResult) ) 

#define IWndPage_AddObjToHtml(This,strObjName,bConnectEvent,pObjDisp)	\
    ( (This)->lpVtbl -> AddObjToHtml(This,strObjName,bConnectEvent,pObjDisp) ) 

#define IWndPage_AddNodesToPage(This,pHtmlDoc,bstrNodeNames,bAdd,bSuccess)	\
    ( (This)->lpVtbl -> AddNodesToPage(This,pHtmlDoc,bstrNodeNames,bAdd,bSuccess) ) 

#define IWndPage_AttachObjEvent(This,HTMLWindow,SourceObj,bstrEventName,AliasName)	\
    ( (This)->lpVtbl -> AttachObjEvent(This,HTMLWindow,SourceObj,bstrEventName,AliasName) ) 

#define IWndPage_AttachNodeSubCtrlEvent(This,HtmlWndDisp,Node,Ctrl,EventName,AliasName)	\
    ( (This)->lpVtbl -> AttachNodeSubCtrlEvent(This,HtmlWndDisp,Node,Ctrl,EventName,AliasName) ) 

#define IWndPage_AttachEvent(This,bstrNames,pHtmlWnd,bResult)	\
    ( (This)->lpVtbl -> AttachEvent(This,bstrNames,pHtmlWnd,bResult) ) 

#define IWndPage_AttachNodeEvent(This,bstrNames,pWndDisp)	\
    ( (This)->lpVtbl -> AttachNodeEvent(This,bstrNames,pWndDisp) ) 

#define IWndPage_Extend(This,Parent,CtrlName,FrameName,bstrKey,bstrXml,ppRetNode)	\
    ( (This)->lpVtbl -> Extend(This,Parent,CtrlName,FrameName,bstrKey,bstrXml,ppRetNode) ) 

#define IWndPage_ExtendCtrl(This,Ctrl,bstrKey,bstrXml,ppRetNode)	\
    ( (This)->lpVtbl -> ExtendCtrl(This,Ctrl,bstrKey,bstrXml,ppRetNode) ) 

#define IWndPage_GetCtrlByName(This,pCtrl,bstrName,bFindInChild,ppCtrlDisp)	\
    ( (This)->lpVtbl -> GetCtrlByName(This,pCtrl,bstrName,bFindInChild,ppCtrlDisp) ) 

#define IWndPage_GetCtrlInNode(This,NodeName,CtrlName,ppCtrl)	\
    ( (This)->lpVtbl -> GetCtrlInNode(This,NodeName,CtrlName,ppCtrl) ) 

#define IWndPage_GetHtmlDocument(This,HtmlWindow,ppHtmlDoc)	\
    ( (This)->lpVtbl -> GetHtmlDocument(This,HtmlWindow,ppHtmlDoc) ) 

#define IWndPage_GetFrameFromCtrl(This,ctrl,ppFrame)	\
    ( (This)->lpVtbl -> GetFrameFromCtrl(This,ctrl,ppFrame) ) 

#define IWndPage_GetWndNode(This,bstrFrameName,bstrNodeName,pRetNode)	\
    ( (This)->lpVtbl -> GetWndNode(This,bstrFrameName,bstrNodeName,pRetNode) ) 

#define IWndPage_GetCtrlValueByName(This,pCtrl,bstrName,bFindInChild,bstrVal)	\
    ( (This)->lpVtbl -> GetCtrlValueByName(This,pCtrl,bstrName,bFindInChild,bstrVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IWndPage_INTERFACE_DEFINED__ */


#ifndef __IEclipseTopWnd_INTERFACE_DEFINED__
#define __IEclipseTopWnd_INTERFACE_DEFINED__

/* interface IEclipseTopWnd */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IEclipseTopWnd;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119825D34")
    IEclipseTopWnd : public IDispatch
    {
    public:
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Handle( 
            /* [retval][out] */ long *pVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SWTExtend( 
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppNode) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetCtrlText( 
            BSTR bstrNodeName,
            BSTR bstrCtrlName,
            /* [retval][out] */ BSTR *bstrVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IEclipseTopWndVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IEclipseTopWnd * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IEclipseTopWnd * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IEclipseTopWnd * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IEclipseTopWnd * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IEclipseTopWnd * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IEclipseTopWnd * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IEclipseTopWnd * This,
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
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Handle )( 
            IEclipseTopWnd * This,
            /* [retval][out] */ long *pVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *SWTExtend )( 
            IEclipseTopWnd * This,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppNode);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetCtrlText )( 
            IEclipseTopWnd * This,
            BSTR bstrNodeName,
            BSTR bstrCtrlName,
            /* [retval][out] */ BSTR *bstrVal);
        
        END_INTERFACE
    } IEclipseTopWndVtbl;

    interface IEclipseTopWnd
    {
        CONST_VTBL struct IEclipseTopWndVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IEclipseTopWnd_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IEclipseTopWnd_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IEclipseTopWnd_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IEclipseTopWnd_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IEclipseTopWnd_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IEclipseTopWnd_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IEclipseTopWnd_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IEclipseTopWnd_get_Handle(This,pVal)	\
    ( (This)->lpVtbl -> get_Handle(This,pVal) ) 

#define IEclipseTopWnd_SWTExtend(This,bstrKey,bstrXml,ppNode)	\
    ( (This)->lpVtbl -> SWTExtend(This,bstrKey,bstrXml,ppNode) ) 

#define IEclipseTopWnd_GetCtrlText(This,bstrNodeName,bstrCtrlName,bstrVal)	\
    ( (This)->lpVtbl -> GetCtrlText(This,bstrNodeName,bstrCtrlName,bstrVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IEclipseTopWnd_INTERFACE_DEFINED__ */


#ifndef __IOfficeObject_INTERFACE_DEFINED__
#define __IOfficeObject_INTERFACE_DEFINED__

/* interface IOfficeObject */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IOfficeObject;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119827857")
    IOfficeObject : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Show( 
            VARIANT_BOOL bShow) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Extend( 
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppNode) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE UnloadTangram( void) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IOfficeObjectVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IOfficeObject * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IOfficeObject * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IOfficeObject * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IOfficeObject * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IOfficeObject * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IOfficeObject * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IOfficeObject * This,
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
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Show )( 
            IOfficeObject * This,
            VARIANT_BOOL bShow);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Extend )( 
            IOfficeObject * This,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppNode);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *UnloadTangram )( 
            IOfficeObject * This);
        
        END_INTERFACE
    } IOfficeObjectVtbl;

    interface IOfficeObject
    {
        CONST_VTBL struct IOfficeObjectVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOfficeObject_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IOfficeObject_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IOfficeObject_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IOfficeObject_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IOfficeObject_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IOfficeObject_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IOfficeObject_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IOfficeObject_Show(This,bShow)	\
    ( (This)->lpVtbl -> Show(This,bShow) ) 

#define IOfficeObject_Extend(This,bstrKey,bstrXml,ppNode)	\
    ( (This)->lpVtbl -> Extend(This,bstrKey,bstrXml,ppNode) ) 

#define IOfficeObject_UnloadTangram(This)	\
    ( (This)->lpVtbl -> UnloadTangram(This) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IOfficeObject_INTERFACE_DEFINED__ */


#ifndef __ITaskObj_INTERFACE_DEFINED__
#define __ITaskObj_INTERFACE_DEFINED__

/* interface ITaskObj */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITaskObj;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119827B88")
    ITaskObj : public IDispatch
    {
    public:
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_TaskXML( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_TaskXML( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Execute( 
            ITaskNotify *pCallBack,
            IDispatch *pStateDisp,
            /* [retval][out] */ IDispatch **DispRet) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateNode( 
            TaskNodeType NodeType,
            BSTR NodeName,
            BSTR bstrXml) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateNode2( 
            TaskNodeType NodeType,
            BSTR NodeName,
            ITaskObj *pTangramTaskObj) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_TaskParticipantObj( 
            BSTR bstrID,
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_TaskParticipantObj( 
            BSTR bstrID,
            /* [in] */ IDispatch *newVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ITaskObjVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITaskObj * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITaskObj * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITaskObj * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITaskObj * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITaskObj * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITaskObj * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITaskObj * This,
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
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_TaskXML )( 
            ITaskObj * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_TaskXML )( 
            ITaskObj * This,
            /* [in] */ BSTR newVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Execute )( 
            ITaskObj * This,
            ITaskNotify *pCallBack,
            IDispatch *pStateDisp,
            /* [retval][out] */ IDispatch **DispRet);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CreateNode )( 
            ITaskObj * This,
            TaskNodeType NodeType,
            BSTR NodeName,
            BSTR bstrXml);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CreateNode2 )( 
            ITaskObj * This,
            TaskNodeType NodeType,
            BSTR NodeName,
            ITaskObj *pTangramTaskObj);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_TaskParticipantObj )( 
            ITaskObj * This,
            BSTR bstrID,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_TaskParticipantObj )( 
            ITaskObj * This,
            BSTR bstrID,
            /* [in] */ IDispatch *newVal);
        
        END_INTERFACE
    } ITaskObjVtbl;

    interface ITaskObj
    {
        CONST_VTBL struct ITaskObjVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITaskObj_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITaskObj_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITaskObj_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITaskObj_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITaskObj_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITaskObj_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITaskObj_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITaskObj_get_TaskXML(This,pVal)	\
    ( (This)->lpVtbl -> get_TaskXML(This,pVal) ) 

#define ITaskObj_put_TaskXML(This,newVal)	\
    ( (This)->lpVtbl -> put_TaskXML(This,newVal) ) 

#define ITaskObj_Execute(This,pCallBack,pStateDisp,DispRet)	\
    ( (This)->lpVtbl -> Execute(This,pCallBack,pStateDisp,DispRet) ) 

#define ITaskObj_CreateNode(This,NodeType,NodeName,bstrXml)	\
    ( (This)->lpVtbl -> CreateNode(This,NodeType,NodeName,bstrXml) ) 

#define ITaskObj_CreateNode2(This,NodeType,NodeName,pTangramTaskObj)	\
    ( (This)->lpVtbl -> CreateNode2(This,NodeType,NodeName,pTangramTaskObj) ) 

#define ITaskObj_get_TaskParticipantObj(This,bstrID,pVal)	\
    ( (This)->lpVtbl -> get_TaskParticipantObj(This,bstrID,pVal) ) 

#define ITaskObj_put_TaskParticipantObj(This,bstrID,newVal)	\
    ( (This)->lpVtbl -> put_TaskParticipantObj(This,bstrID,newVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITaskObj_INTERFACE_DEFINED__ */


#ifndef __IWebTaskObj_INTERFACE_DEFINED__
#define __IWebTaskObj_INTERFACE_DEFINED__

/* interface IWebTaskObj */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IWebTaskObj;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-060119825430")
    IWebTaskObj : public IDispatch
    {
    public:
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_XML( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_XML( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Type( 
            /* [retval][out] */ int *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Type( 
            /* [in] */ int newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Extender( 
            BSTR bstrExtenderName,
            /* [retval][out] */ IDispatch **pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Extender( 
            BSTR bstrExtenderName,
            /* [in] */ IDispatch *newVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Run( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Quit( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE InitWebConnection( 
            BSTR bstrURL,
            ITaskNotify *pDispNotify) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE TangramAction( 
            BSTR bstrXml,
            ITaskNotify *pDispNotify) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE execScript( 
            BSTR bstrScript,
            /* [retval][out] */ VARIANT *varRet) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IWebTaskObjVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IWebTaskObj * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IWebTaskObj * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IWebTaskObj * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IWebTaskObj * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IWebTaskObj * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IWebTaskObj * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IWebTaskObj * This,
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
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_XML )( 
            IWebTaskObj * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_XML )( 
            IWebTaskObj * This,
            /* [in] */ BSTR newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Type )( 
            IWebTaskObj * This,
            /* [retval][out] */ int *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Type )( 
            IWebTaskObj * This,
            /* [in] */ int newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Extender )( 
            IWebTaskObj * This,
            BSTR bstrExtenderName,
            /* [retval][out] */ IDispatch **pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Extender )( 
            IWebTaskObj * This,
            BSTR bstrExtenderName,
            /* [in] */ IDispatch *newVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Run )( 
            IWebTaskObj * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Quit )( 
            IWebTaskObj * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *InitWebConnection )( 
            IWebTaskObj * This,
            BSTR bstrURL,
            ITaskNotify *pDispNotify);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *TangramAction )( 
            IWebTaskObj * This,
            BSTR bstrXml,
            ITaskNotify *pDispNotify);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *execScript )( 
            IWebTaskObj * This,
            BSTR bstrScript,
            /* [retval][out] */ VARIANT *varRet);
        
        END_INTERFACE
    } IWebTaskObjVtbl;

    interface IWebTaskObj
    {
        CONST_VTBL struct IWebTaskObjVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWebTaskObj_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IWebTaskObj_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IWebTaskObj_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IWebTaskObj_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IWebTaskObj_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IWebTaskObj_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IWebTaskObj_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IWebTaskObj_get_XML(This,pVal)	\
    ( (This)->lpVtbl -> get_XML(This,pVal) ) 

#define IWebTaskObj_put_XML(This,newVal)	\
    ( (This)->lpVtbl -> put_XML(This,newVal) ) 

#define IWebTaskObj_get_Type(This,pVal)	\
    ( (This)->lpVtbl -> get_Type(This,pVal) ) 

#define IWebTaskObj_put_Type(This,newVal)	\
    ( (This)->lpVtbl -> put_Type(This,newVal) ) 

#define IWebTaskObj_get_Extender(This,bstrExtenderName,pVal)	\
    ( (This)->lpVtbl -> get_Extender(This,bstrExtenderName,pVal) ) 

#define IWebTaskObj_put_Extender(This,bstrExtenderName,newVal)	\
    ( (This)->lpVtbl -> put_Extender(This,bstrExtenderName,newVal) ) 

#define IWebTaskObj_Run(This)	\
    ( (This)->lpVtbl -> Run(This) ) 

#define IWebTaskObj_Quit(This)	\
    ( (This)->lpVtbl -> Quit(This) ) 

#define IWebTaskObj_InitWebConnection(This,bstrURL,pDispNotify)	\
    ( (This)->lpVtbl -> InitWebConnection(This,bstrURL,pDispNotify) ) 

#define IWebTaskObj_TangramAction(This,bstrXml,pDispNotify)	\
    ( (This)->lpVtbl -> TangramAction(This,bstrXml,pDispNotify) ) 

#define IWebTaskObj_execScript(This,bstrScript,varRet)	\
    ( (This)->lpVtbl -> execScript(This,bstrScript,varRet) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IWebTaskObj_INTERFACE_DEFINED__ */


#ifndef __IRestObject_INTERFACE_DEFINED__
#define __IRestObject_INTERFACE_DEFINED__

/* interface IRestObject */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IRestObject;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-06011982121C")
    IRestObject : public IDispatch
    {
    public:
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CloudAddinRestNotify( 
            /* [retval][out] */ IRestNotify **pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_CloudAddinRestNotify( 
            /* [in] */ IRestNotify *newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_NotifyHandle( 
            /* [retval][out] */ LONGLONG *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_NotifyHandle( 
            /* [in] */ LONGLONG newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Header( 
            BSTR bstrHeaderName,
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Header( 
            BSTR bstrHeaderName,
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_RestKey( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_RestKey( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_WndNode( 
            /* [retval][out] */ IWndNode **pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_WndNode( 
            /* [in] */ IWndNode *newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_State( 
            /* [retval][out] */ int *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_State( 
            /* [in] */ int newVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE RestData( 
            int nMethod,
            BSTR bstrServerURL,
            BSTR bstrRequest,
            BSTR bstrData,
            LONGLONG hNotify) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ClearHeaders( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clone( 
            IRestObject *pTargetObj) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AsyncRest( 
            int nMethod,
            BSTR bstrURL,
            BSTR bstrData,
            LONGLONG hNotify) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Notify( 
            long nNotify) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE UploadFile( 
            VARIANT_BOOL bUpload,
            BSTR bstrServerURL,
            BSTR bstrRequest,
            BSTR bstrFilePath) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IRestObjectVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IRestObject * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IRestObject * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IRestObject * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IRestObject * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IRestObject * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IRestObject * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IRestObject * This,
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
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_CloudAddinRestNotify )( 
            IRestObject * This,
            /* [retval][out] */ IRestNotify **pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_CloudAddinRestNotify )( 
            IRestObject * This,
            /* [in] */ IRestNotify *newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_NotifyHandle )( 
            IRestObject * This,
            /* [retval][out] */ LONGLONG *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_NotifyHandle )( 
            IRestObject * This,
            /* [in] */ LONGLONG newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Header )( 
            IRestObject * This,
            BSTR bstrHeaderName,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Header )( 
            IRestObject * This,
            BSTR bstrHeaderName,
            /* [in] */ BSTR newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_RestKey )( 
            IRestObject * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_RestKey )( 
            IRestObject * This,
            /* [in] */ BSTR newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_WndNode )( 
            IRestObject * This,
            /* [retval][out] */ IWndNode **pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_WndNode )( 
            IRestObject * This,
            /* [in] */ IWndNode *newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_State )( 
            IRestObject * This,
            /* [retval][out] */ int *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_State )( 
            IRestObject * This,
            /* [in] */ int newVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *RestData )( 
            IRestObject * This,
            int nMethod,
            BSTR bstrServerURL,
            BSTR bstrRequest,
            BSTR bstrData,
            LONGLONG hNotify);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ClearHeaders )( 
            IRestObject * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Clone )( 
            IRestObject * This,
            IRestObject *pTargetObj);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AsyncRest )( 
            IRestObject * This,
            int nMethod,
            BSTR bstrURL,
            BSTR bstrData,
            LONGLONG hNotify);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Notify )( 
            IRestObject * This,
            long nNotify);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *UploadFile )( 
            IRestObject * This,
            VARIANT_BOOL bUpload,
            BSTR bstrServerURL,
            BSTR bstrRequest,
            BSTR bstrFilePath);
        
        END_INTERFACE
    } IRestObjectVtbl;

    interface IRestObject
    {
        CONST_VTBL struct IRestObjectVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IRestObject_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IRestObject_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IRestObject_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IRestObject_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IRestObject_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IRestObject_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IRestObject_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IRestObject_get_CloudAddinRestNotify(This,pVal)	\
    ( (This)->lpVtbl -> get_CloudAddinRestNotify(This,pVal) ) 

#define IRestObject_put_CloudAddinRestNotify(This,newVal)	\
    ( (This)->lpVtbl -> put_CloudAddinRestNotify(This,newVal) ) 

#define IRestObject_get_NotifyHandle(This,pVal)	\
    ( (This)->lpVtbl -> get_NotifyHandle(This,pVal) ) 

#define IRestObject_put_NotifyHandle(This,newVal)	\
    ( (This)->lpVtbl -> put_NotifyHandle(This,newVal) ) 

#define IRestObject_get_Header(This,bstrHeaderName,pVal)	\
    ( (This)->lpVtbl -> get_Header(This,bstrHeaderName,pVal) ) 

#define IRestObject_put_Header(This,bstrHeaderName,newVal)	\
    ( (This)->lpVtbl -> put_Header(This,bstrHeaderName,newVal) ) 

#define IRestObject_get_RestKey(This,pVal)	\
    ( (This)->lpVtbl -> get_RestKey(This,pVal) ) 

#define IRestObject_put_RestKey(This,newVal)	\
    ( (This)->lpVtbl -> put_RestKey(This,newVal) ) 

#define IRestObject_get_WndNode(This,pVal)	\
    ( (This)->lpVtbl -> get_WndNode(This,pVal) ) 

#define IRestObject_put_WndNode(This,newVal)	\
    ( (This)->lpVtbl -> put_WndNode(This,newVal) ) 

#define IRestObject_get_State(This,pVal)	\
    ( (This)->lpVtbl -> get_State(This,pVal) ) 

#define IRestObject_put_State(This,newVal)	\
    ( (This)->lpVtbl -> put_State(This,newVal) ) 

#define IRestObject_RestData(This,nMethod,bstrServerURL,bstrRequest,bstrData,hNotify)	\
    ( (This)->lpVtbl -> RestData(This,nMethod,bstrServerURL,bstrRequest,bstrData,hNotify) ) 

#define IRestObject_ClearHeaders(This)	\
    ( (This)->lpVtbl -> ClearHeaders(This) ) 

#define IRestObject_Clone(This,pTargetObj)	\
    ( (This)->lpVtbl -> Clone(This,pTargetObj) ) 

#define IRestObject_AsyncRest(This,nMethod,bstrURL,bstrData,hNotify)	\
    ( (This)->lpVtbl -> AsyncRest(This,nMethod,bstrURL,bstrData,hNotify) ) 

#define IRestObject_Notify(This,nNotify)	\
    ( (This)->lpVtbl -> Notify(This,nNotify) ) 

#define IRestObject_UploadFile(This,bUpload,bstrServerURL,bstrRequest,bstrFilePath)	\
    ( (This)->lpVtbl -> UploadFile(This,bUpload,bstrServerURL,bstrRequest,bstrFilePath) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IRestObject_INTERFACE_DEFINED__ */


#ifndef __ITangramBot_INTERFACE_DEFINED__
#define __ITangramBot_INTERFACE_DEFINED__

/* interface ITangramBot */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITangramBot;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("EC57580E-2223-48E5-B9AA-5D67AB18EBF1")
    ITangramBot : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct ITangramBotVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITangramBot * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITangramBot * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITangramBot * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITangramBot * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITangramBot * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITangramBot * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITangramBot * This,
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
    } ITangramBotVtbl;

    interface ITangramBot
    {
        CONST_VTBL struct ITangramBotVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITangramBot_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITangramBot_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITangramBot_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITangramBot_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITangramBot_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITangramBot_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITangramBot_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITangramBot_INTERFACE_DEFINED__ */


#ifndef __ICollaborationProxy_INTERFACE_DEFINED__
#define __ICollaborationProxy_INTERFACE_DEFINED__

/* interface ICollaborationProxy */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICollaborationProxy;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0882EEBD-559F-4A10-ABA9-D526BF321B1D")
    ICollaborationProxy : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Start( 
            BSTR bstrXml) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Close( 
            BSTR bstrXml) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICollaborationProxyVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICollaborationProxy * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICollaborationProxy * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICollaborationProxy * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICollaborationProxy * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICollaborationProxy * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICollaborationProxy * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICollaborationProxy * This,
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
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Start )( 
            ICollaborationProxy * This,
            BSTR bstrXml);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Close )( 
            ICollaborationProxy * This,
            BSTR bstrXml);
        
        END_INTERFACE
    } ICollaborationProxyVtbl;

    interface ICollaborationProxy
    {
        CONST_VTBL struct ICollaborationProxyVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICollaborationProxy_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICollaborationProxy_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICollaborationProxy_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICollaborationProxy_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICollaborationProxy_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICollaborationProxy_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICollaborationProxy_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICollaborationProxy_Start(This,bstrXml)	\
    ( (This)->lpVtbl -> Start(This,bstrXml) ) 

#define ICollaborationProxy_Close(This,bstrXml)	\
    ( (This)->lpVtbl -> Close(This,bstrXml) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICollaborationProxy_INTERFACE_DEFINED__ */


#ifndef __ITangramCtrl_INTERFACE_DEFINED__
#define __ITangramCtrl_INTERFACE_DEFINED__

/* interface ITangramCtrl */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITangramCtrl;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("19631222-1992-0612-1965-0601198231DC")
    ITangramCtrl : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_HWND( 
            /* [retval][out] */ long *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_TangramHandle( 
            BSTR bstrHandleName,
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetCtrlText( 
            long nHandle,
            BSTR bstrNodeName,
            BSTR bstrCtrlName,
            /* [retval][out] */ BSTR *bstrVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE InitCtrl( 
            BSTR bstrXml) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Extend( 
            BSTR bstrFrameName,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppNode) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ITangramCtrlVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITangramCtrl * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITangramCtrl * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITangramCtrl * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITangramCtrl * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITangramCtrl * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITangramCtrl * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITangramCtrl * This,
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
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HWND )( 
            ITangramCtrl * This,
            /* [retval][out] */ long *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_TangramHandle )( 
            ITangramCtrl * This,
            BSTR bstrHandleName,
            /* [in] */ LONG newVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetCtrlText )( 
            ITangramCtrl * This,
            long nHandle,
            BSTR bstrNodeName,
            BSTR bstrCtrlName,
            /* [retval][out] */ BSTR *bstrVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *InitCtrl )( 
            ITangramCtrl * This,
            BSTR bstrXml);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Extend )( 
            ITangramCtrl * This,
            BSTR bstrFrameName,
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppNode);
        
        END_INTERFACE
    } ITangramCtrlVtbl;

    interface ITangramCtrl
    {
        CONST_VTBL struct ITangramCtrlVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITangramCtrl_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITangramCtrl_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITangramCtrl_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITangramCtrl_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITangramCtrl_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITangramCtrl_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITangramCtrl_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITangramCtrl_get_HWND(This,pVal)	\
    ( (This)->lpVtbl -> get_HWND(This,pVal) ) 

#define ITangramCtrl_put_TangramHandle(This,bstrHandleName,newVal)	\
    ( (This)->lpVtbl -> put_TangramHandle(This,bstrHandleName,newVal) ) 

#define ITangramCtrl_GetCtrlText(This,nHandle,bstrNodeName,bstrCtrlName,bstrVal)	\
    ( (This)->lpVtbl -> GetCtrlText(This,nHandle,bstrNodeName,bstrCtrlName,bstrVal) ) 

#define ITangramCtrl_InitCtrl(This,bstrXml)	\
    ( (This)->lpVtbl -> InitCtrl(This,bstrXml) ) 

#define ITangramCtrl_Extend(This,bstrFrameName,bstrKey,bstrXml,ppNode)	\
    ( (This)->lpVtbl -> Extend(This,bstrFrameName,bstrKey,bstrXml,ppNode) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITangramCtrl_INTERFACE_DEFINED__ */



#ifndef __Tangram_LIBRARY_DEFINED__
#define __Tangram_LIBRARY_DEFINED__

/* library Tangram */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_Tangram;

#ifndef ___ITangramEvents_DISPINTERFACE_DEFINED__
#define ___ITangramEvents_DISPINTERFACE_DEFINED__

/* dispinterface _ITangramEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__ITangramEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("19631222-1992-0612-1965-060119821002")
    _ITangramEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _ITangramEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _ITangramEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _ITangramEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _ITangramEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _ITangramEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _ITangramEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _ITangramEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _ITangramEvents * This,
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
    } _ITangramEventsVtbl;

    interface _ITangramEvents
    {
        CONST_VTBL struct _ITangramEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _ITangramEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _ITangramEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _ITangramEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _ITangramEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _ITangramEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _ITangramEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _ITangramEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___ITangramEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_Tangram;

#ifdef __cplusplus

class DECLSPEC_UUID("19631222-1992-0612-1965-060119822002")
Tangram;
#endif

#ifndef ___IWndNodeEvents_DISPINTERFACE_DEFINED__
#define ___IWndNodeEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IWndNodeEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IWndNodeEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("19631222-1992-0612-1965-060119821001")
    _IWndNodeEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IWndNodeEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IWndNodeEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IWndNodeEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IWndNodeEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IWndNodeEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IWndNodeEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IWndNodeEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IWndNodeEvents * This,
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
    } _IWndNodeEventsVtbl;

    interface _IWndNodeEvents
    {
        CONST_VTBL struct _IWndNodeEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IWndNodeEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _IWndNodeEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _IWndNodeEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _IWndNodeEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _IWndNodeEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _IWndNodeEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _IWndNodeEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IWndNodeEvents_DISPINTERFACE_DEFINED__ */


#ifndef ___IWndPage_DISPINTERFACE_DEFINED__
#define ___IWndPage_DISPINTERFACE_DEFINED__

/* dispinterface _IWndPage */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IWndPage;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("19631222-1992-0612-1965-060119821871")
    _IWndPage : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IWndPageVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IWndPage * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IWndPage * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IWndPage * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IWndPage * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IWndPage * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IWndPage * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IWndPage * This,
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
    } _IWndPageVtbl;

    interface _IWndPage
    {
        CONST_VTBL struct _IWndPageVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IWndPage_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _IWndPage_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _IWndPage_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _IWndPage_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _IWndPage_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _IWndPage_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _IWndPage_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IWndPage_DISPINTERFACE_DEFINED__ */


#ifndef ___IEventProxy_DISPINTERFACE_DEFINED__
#define ___IEventProxy_DISPINTERFACE_DEFINED__

/* dispinterface _IEventProxy */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IEventProxy;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("19631222-1992-0612-1965-060119821963")
    _IEventProxy : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IEventProxyVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IEventProxy * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IEventProxy * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IEventProxy * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IEventProxy * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IEventProxy * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IEventProxy * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IEventProxy * This,
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
    } _IEventProxyVtbl;

    interface _IEventProxy
    {
        CONST_VTBL struct _IEventProxyVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IEventProxy_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _IEventProxy_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _IEventProxy_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _IEventProxy_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _IEventProxy_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _IEventProxy_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _IEventProxy_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IEventProxy_DISPINTERFACE_DEFINED__ */


#ifndef ___IOfficeObjectEvents_DISPINTERFACE_DEFINED__
#define ___IOfficeObjectEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IOfficeObjectEvents */
/* [uuid] */ 


EXTERN_C const IID DIID__IOfficeObjectEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("19631222-1992-0612-1965-06011982D24E")
    _IOfficeObjectEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IOfficeObjectEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IOfficeObjectEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IOfficeObjectEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IOfficeObjectEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IOfficeObjectEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IOfficeObjectEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IOfficeObjectEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IOfficeObjectEvents * This,
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
    } _IOfficeObjectEventsVtbl;

    interface _IOfficeObjectEvents
    {
        CONST_VTBL struct _IOfficeObjectEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IOfficeObjectEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _IOfficeObjectEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _IOfficeObjectEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _IOfficeObjectEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _IOfficeObjectEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _IOfficeObjectEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _IOfficeObjectEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IOfficeObjectEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_OfficeObject;

#ifdef __cplusplus

class DECLSPEC_UUID("19631222-1992-0612-1965-060119827289")
OfficeObject;
#endif

#ifndef ___ITangramBotEvents_DISPINTERFACE_DEFINED__
#define ___ITangramBotEvents_DISPINTERFACE_DEFINED__

/* dispinterface _ITangramBotEvents */
/* [uuid] */ 


EXTERN_C const IID DIID__ITangramBotEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("98625504-E0E7-43CC-928D-CA20DF398525")
    _ITangramBotEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _ITangramBotEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _ITangramBotEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _ITangramBotEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _ITangramBotEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _ITangramBotEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _ITangramBotEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _ITangramBotEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _ITangramBotEvents * This,
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
    } _ITangramBotEventsVtbl;

    interface _ITangramBotEvents
    {
        CONST_VTBL struct _ITangramBotEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _ITangramBotEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _ITangramBotEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _ITangramBotEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _ITangramBotEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _ITangramBotEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _ITangramBotEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _ITangramBotEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___ITangramBotEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_TangramBot;

#ifdef __cplusplus

class DECLSPEC_UUID("CD8F9CD1-32A2-40B3-BCCA-CEF249A0DFF2")
TangramBot;
#endif

EXTERN_C const CLSID CLSID_CollaborationProxy;

#ifdef __cplusplus

class DECLSPEC_UUID("4F44F00D-6E43-4023-A71D-C3AF7618DF10")
CollaborationProxy;
#endif

EXTERN_C const CLSID CLSID_TangramCtrl;

#ifdef __cplusplus

class DECLSPEC_UUID("19631222-1992-0612-1965-060119821986")
TangramCtrl;
#endif
#endif /* __Tangram_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  VARIANT_UserSize(     unsigned long *, unsigned long            , VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserMarshal(  unsigned long *, unsigned char *, VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserUnmarshal(unsigned long *, unsigned char *, VARIANT * ); 
void                      __RPC_USER  VARIANT_UserFree(     unsigned long *, VARIANT * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


