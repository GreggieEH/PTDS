

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Tue Sep 06 12:12:03 2016
 */
/* Compiler settings for MyIDL.idl:
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


#ifndef __MyIDL_h_h__
#define __MyIDL_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IPTDS_FWD_DEFINED__
#define __IPTDS_FWD_DEFINED__
typedef interface IPTDS IPTDS;

#endif 	/* __IPTDS_FWD_DEFINED__ */


#ifndef ___PTDS_FWD_DEFINED__
#define ___PTDS_FWD_DEFINED__
typedef interface _PTDS _PTDS;

#endif 	/* ___PTDS_FWD_DEFINED__ */


#ifndef __PTDS_FWD_DEFINED__
#define __PTDS_FWD_DEFINED__

#ifdef __cplusplus
typedef class PTDS PTDS;
#else
typedef struct PTDS PTDS;
#endif /* __cplusplus */

#endif 	/* __PTDS_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_MyIDL_0000_0000 */
/* [local] */ 

#pragma once


extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_s_ifspec;


#ifndef __PTDS_LIBRARY_DEFINED__
#define __PTDS_LIBRARY_DEFINED__

/* library PTDS */
/* [version][helpstring][uuid] */ 


EXTERN_C const IID LIBID_PTDS;

#ifndef __IPTDS_DISPINTERFACE_DEFINED__
#define __IPTDS_DISPINTERFACE_DEFINED__

/* dispinterface IPTDS */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IPTDS;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("CCEA4795-53CF-4A91-99CE-19DF217BD322")
    IPTDS : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IPTDSVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IPTDS * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IPTDS * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IPTDS * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IPTDS * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IPTDS * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IPTDS * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IPTDS * This,
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
    } IPTDSVtbl;

    interface IPTDS
    {
        CONST_VTBL struct IPTDSVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IPTDS_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IPTDS_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IPTDS_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IPTDS_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IPTDS_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IPTDS_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IPTDS_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IPTDS_DISPINTERFACE_DEFINED__ */


#ifndef ___PTDS_DISPINTERFACE_DEFINED__
#define ___PTDS_DISPINTERFACE_DEFINED__

/* dispinterface _PTDS */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__PTDS;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("D56900BB-5B79-468F-BE4D-023413ACAC39")
    _PTDS : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _PTDSVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _PTDS * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _PTDS * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _PTDS * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _PTDS * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _PTDS * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _PTDS * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _PTDS * This,
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
    } _PTDSVtbl;

    interface _PTDS
    {
        CONST_VTBL struct _PTDSVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _PTDS_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _PTDS_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _PTDS_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _PTDS_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _PTDS_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _PTDS_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _PTDS_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___PTDS_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_PTDS;

#ifdef __cplusplus

class DECLSPEC_UUID("7AABE1F0-260D-4FE8-8657-6FA40A818986")
PTDS;
#endif
#endif /* __PTDS_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


