

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 6.00.0361 */
/* at Thu Feb 24 13:17:42 2005
 */
/* Compiler settings for .\gen_activewa.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

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

#ifndef __gen_activewa_h__
#define __gen_activewa_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IPlaylist_FWD_DEFINED__
#define __IPlaylist_FWD_DEFINED__
typedef interface IPlaylist IPlaylist;
#endif 	/* __IPlaylist_FWD_DEFINED__ */


#ifndef __IApplication_FWD_DEFINED__
#define __IApplication_FWD_DEFINED__
typedef interface IApplication IApplication;
#endif 	/* __IApplication_FWD_DEFINED__ */


#ifndef __IMediaItem_FWD_DEFINED__
#define __IMediaItem_FWD_DEFINED__
typedef interface IMediaItem IMediaItem;
#endif 	/* __IMediaItem_FWD_DEFINED__ */


#ifndef __IMediaLibrary_FWD_DEFINED__
#define __IMediaLibrary_FWD_DEFINED__
typedef interface IMediaLibrary IMediaLibrary;
#endif 	/* __IMediaLibrary_FWD_DEFINED__ */


#ifndef __ISiteManager_FWD_DEFINED__
#define __ISiteManager_FWD_DEFINED__
typedef interface ISiteManager ISiteManager;
#endif 	/* __ISiteManager_FWD_DEFINED__ */


#ifndef ___IApplicationEvents_FWD_DEFINED__
#define ___IApplicationEvents_FWD_DEFINED__
typedef interface _IApplicationEvents _IApplicationEvents;
#endif 	/* ___IApplicationEvents_FWD_DEFINED__ */


#ifndef __Application_FWD_DEFINED__
#define __Application_FWD_DEFINED__

#ifdef __cplusplus
typedef class Application Application;
#else
typedef struct Application Application;
#endif /* __cplusplus */

#endif 	/* __Application_FWD_DEFINED__ */


#ifndef __Playlist_FWD_DEFINED__
#define __Playlist_FWD_DEFINED__

#ifdef __cplusplus
typedef class Playlist Playlist;
#else
typedef struct Playlist Playlist;
#endif /* __cplusplus */

#endif 	/* __Playlist_FWD_DEFINED__ */


#ifndef __MediaItem_FWD_DEFINED__
#define __MediaItem_FWD_DEFINED__

#ifdef __cplusplus
typedef class MediaItem MediaItem;
#else
typedef struct MediaItem MediaItem;
#endif /* __cplusplus */

#endif 	/* __MediaItem_FWD_DEFINED__ */


#ifndef __MediaLibrary_FWD_DEFINED__
#define __MediaLibrary_FWD_DEFINED__

#ifdef __cplusplus
typedef class MediaLibrary MediaLibrary;
#else
typedef struct MediaLibrary MediaLibrary;
#endif /* __cplusplus */

#endif 	/* __MediaLibrary_FWD_DEFINED__ */


#ifndef __SiteManager_FWD_DEFINED__
#define __SiteManager_FWD_DEFINED__

#ifdef __cplusplus
typedef class SiteManager SiteManager;
#else
typedef struct SiteManager SiteManager;
#endif /* __cplusplus */

#endif 	/* __SiteManager_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 

void * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void * ); 

/* interface __MIDL_itf_gen_activewa_0000 */
/* [local] */ 





extern RPC_IF_HANDLE __MIDL_itf_gen_activewa_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_gen_activewa_0000_v0_0_s_ifspec;

#ifndef __IPlaylist_INTERFACE_DEFINED__
#define __IPlaylist_INTERFACE_DEFINED__

/* interface IPlaylist */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IPlaylist;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("83F2DBE6-FA22-4D6B-A281-0427D7AF490C")
    IPlaylist : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][restricted][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ LONG index,
            /* [retval][out] */ IMediaItem **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Position( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Position( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetSelection( 
            /* [retval][out] */ VARIANT *SelectionArray) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE DeleteIndex( 
            /* [in] */ LONG index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SwapIndex( 
            /* [in] */ INT from,
            /* [in] */ INT to) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE FlushCache( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Hwnd( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SendMsg( 
            /* [in] */ LONG msg,
            /* [in] */ LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE PostMsg( 
            /* [in] */ LONG msg,
            LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IPlaylistVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IPlaylist * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IPlaylist * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IPlaylist * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IPlaylist * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IPlaylist * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IPlaylist * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IPlaylist * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Count )( 
            IPlaylist * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][restricted][propget] */ HRESULT ( STDMETHODCALLTYPE *get__NewEnum )( 
            IPlaylist * This,
            /* [retval][out] */ IUnknown **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Item )( 
            IPlaylist * This,
            /* [in] */ LONG index,
            /* [retval][out] */ IMediaItem **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Position )( 
            IPlaylist * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Position )( 
            IPlaylist * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetSelection )( 
            IPlaylist * This,
            /* [retval][out] */ VARIANT *SelectionArray);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Clear )( 
            IPlaylist * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *DeleteIndex )( 
            IPlaylist * This,
            /* [in] */ LONG index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SwapIndex )( 
            IPlaylist * This,
            /* [in] */ INT from,
            /* [in] */ INT to);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *FlushCache )( 
            IPlaylist * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Hwnd )( 
            IPlaylist * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SendMsg )( 
            IPlaylist * This,
            /* [in] */ LONG msg,
            /* [in] */ LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *PostMsg )( 
            IPlaylist * This,
            /* [in] */ LONG msg,
            LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result);
        
        END_INTERFACE
    } IPlaylistVtbl;

    interface IPlaylist
    {
        CONST_VTBL struct IPlaylistVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IPlaylist_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IPlaylist_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IPlaylist_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IPlaylist_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IPlaylist_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IPlaylist_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IPlaylist_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IPlaylist_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define IPlaylist_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define IPlaylist_get_Item(This,index,pVal)	\
    (This)->lpVtbl -> get_Item(This,index,pVal)

#define IPlaylist_get_Position(This,pVal)	\
    (This)->lpVtbl -> get_Position(This,pVal)

#define IPlaylist_put_Position(This,newVal)	\
    (This)->lpVtbl -> put_Position(This,newVal)

#define IPlaylist_GetSelection(This,SelectionArray)	\
    (This)->lpVtbl -> GetSelection(This,SelectionArray)

#define IPlaylist_Clear(This)	\
    (This)->lpVtbl -> Clear(This)

#define IPlaylist_DeleteIndex(This,index)	\
    (This)->lpVtbl -> DeleteIndex(This,index)

#define IPlaylist_SwapIndex(This,from,to)	\
    (This)->lpVtbl -> SwapIndex(This,from,to)

#define IPlaylist_FlushCache(This)	\
    (This)->lpVtbl -> FlushCache(This)

#define IPlaylist_get_Hwnd(This,pVal)	\
    (This)->lpVtbl -> get_Hwnd(This,pVal)

#define IPlaylist_SendMsg(This,msg,wParam,lParam,result)	\
    (This)->lpVtbl -> SendMsg(This,msg,wParam,lParam,result)

#define IPlaylist_PostMsg(This,msg,wParam,lParam,result)	\
    (This)->lpVtbl -> PostMsg(This,msg,wParam,lParam,result)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IPlaylist_get_Count_Proxy( 
    IPlaylist * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IPlaylist_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][restricted][propget] */ HRESULT STDMETHODCALLTYPE IPlaylist_get__NewEnum_Proxy( 
    IPlaylist * This,
    /* [retval][out] */ IUnknown **pVal);


void __RPC_STUB IPlaylist_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IPlaylist_get_Item_Proxy( 
    IPlaylist * This,
    /* [in] */ LONG index,
    /* [retval][out] */ IMediaItem **pVal);


void __RPC_STUB IPlaylist_get_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IPlaylist_get_Position_Proxy( 
    IPlaylist * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IPlaylist_get_Position_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IPlaylist_put_Position_Proxy( 
    IPlaylist * This,
    /* [in] */ LONG newVal);


void __RPC_STUB IPlaylist_put_Position_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IPlaylist_GetSelection_Proxy( 
    IPlaylist * This,
    /* [retval][out] */ VARIANT *SelectionArray);


void __RPC_STUB IPlaylist_GetSelection_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IPlaylist_Clear_Proxy( 
    IPlaylist * This);


void __RPC_STUB IPlaylist_Clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IPlaylist_DeleteIndex_Proxy( 
    IPlaylist * This,
    /* [in] */ LONG index);


void __RPC_STUB IPlaylist_DeleteIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IPlaylist_SwapIndex_Proxy( 
    IPlaylist * This,
    /* [in] */ INT from,
    /* [in] */ INT to);


void __RPC_STUB IPlaylist_SwapIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IPlaylist_FlushCache_Proxy( 
    IPlaylist * This);


void __RPC_STUB IPlaylist_FlushCache_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IPlaylist_get_Hwnd_Proxy( 
    IPlaylist * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IPlaylist_get_Hwnd_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IPlaylist_SendMsg_Proxy( 
    IPlaylist * This,
    /* [in] */ LONG msg,
    /* [in] */ LONG wParam,
    /* [in] */ LONGLONG lParam,
    /* [retval][out] */ LONG *result);


void __RPC_STUB IPlaylist_SendMsg_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IPlaylist_PostMsg_Proxy( 
    IPlaylist * This,
    /* [in] */ LONG msg,
    LONG wParam,
    /* [in] */ LONGLONG lParam,
    /* [retval][out] */ LONG *result);


void __RPC_STUB IPlaylist_PostMsg_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IPlaylist_INTERFACE_DEFINED__ */


#ifndef __IApplication_INTERFACE_DEFINED__
#define __IApplication_INTERFACE_DEFINED__

/* interface IApplication */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IApplication;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("2EBD7857-B229-4247-9FAA-17C505410474")
    IApplication : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SayHi( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Playlist( 
            /* [retval][out] */ IPlaylist **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_MediaLibrary( 
            /* [retval][out] */ IMediaLibrary **pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Play( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE StopPlayback( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Pause( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Previous( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Skip( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LoadItem( 
            /* [in] */ BSTR FileName,
            /* [retval][out] */ IMediaItem **MediaItem) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetTimeout( 
            /* [in] */ LONG timeout,
            IDispatch *timeoutfunction,
            /* [retval][out] */ LONG *timerid) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CancelTimer( 
            /* [in] */ LONG timerid) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetIniFile( 
            /* [retval][out] */ BSTR *inifilename) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetIniDirectory( 
            /* [retval][out] */ BSTR *iniDirectory) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ExecVisPlugin( 
            /* [in] */ BSTR VisDllFile) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Skin( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Skin( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Shuffle( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Shuffle( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Repeat( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Repeat( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RestartWinamp( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ShowNotification( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetWaVersion( 
            /* [retval][out] */ LONG *version) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetSendToItems( 
            /* [retval][out] */ VARIANT *Items) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Volume( 
            /* [retval][out] */ BYTE *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Volume( 
            /* [in] */ BYTE newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Panning( 
            /* [retval][out] */ INT *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Panning( 
            /* [in] */ INT newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_PlayState( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Position( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Position( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RunScript( 
            /* [in] */ BSTR scriptfile,
            /* [in] */ BSTR arguments) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE UpdateTitle( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Hwnd( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SendMsg( 
            /* [in] */ LONG msg,
            /* [in] */ LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE PostMsg( 
            /* [in] */ LONG msg,
            /* [in] */ LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IApplicationVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IApplication * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IApplication * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IApplication * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IApplication * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IApplication * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IApplication * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IApplication * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SayHi )( 
            IApplication * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Playlist )( 
            IApplication * This,
            /* [retval][out] */ IPlaylist **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_MediaLibrary )( 
            IApplication * This,
            /* [retval][out] */ IMediaLibrary **pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Play )( 
            IApplication * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *StopPlayback )( 
            IApplication * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Pause )( 
            IApplication * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Previous )( 
            IApplication * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Skip )( 
            IApplication * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *LoadItem )( 
            IApplication * This,
            /* [in] */ BSTR FileName,
            /* [retval][out] */ IMediaItem **MediaItem);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SetTimeout )( 
            IApplication * This,
            /* [in] */ LONG timeout,
            IDispatch *timeoutfunction,
            /* [retval][out] */ LONG *timerid);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CancelTimer )( 
            IApplication * This,
            /* [in] */ LONG timerid);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetIniFile )( 
            IApplication * This,
            /* [retval][out] */ BSTR *inifilename);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetIniDirectory )( 
            IApplication * This,
            /* [retval][out] */ BSTR *iniDirectory);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ExecVisPlugin )( 
            IApplication * This,
            /* [in] */ BSTR VisDllFile);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Skin )( 
            IApplication * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Skin )( 
            IApplication * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Shuffle )( 
            IApplication * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Shuffle )( 
            IApplication * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Repeat )( 
            IApplication * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Repeat )( 
            IApplication * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *RestartWinamp )( 
            IApplication * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ShowNotification )( 
            IApplication * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetWaVersion )( 
            IApplication * This,
            /* [retval][out] */ LONG *version);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetSendToItems )( 
            IApplication * This,
            /* [retval][out] */ VARIANT *Items);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Volume )( 
            IApplication * This,
            /* [retval][out] */ BYTE *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Volume )( 
            IApplication * This,
            /* [in] */ BYTE newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Panning )( 
            IApplication * This,
            /* [retval][out] */ INT *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Panning )( 
            IApplication * This,
            /* [in] */ INT newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_PlayState )( 
            IApplication * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Position )( 
            IApplication * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Position )( 
            IApplication * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *RunScript )( 
            IApplication * This,
            /* [in] */ BSTR scriptfile,
            /* [in] */ BSTR arguments);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *UpdateTitle )( 
            IApplication * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Hwnd )( 
            IApplication * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SendMsg )( 
            IApplication * This,
            /* [in] */ LONG msg,
            /* [in] */ LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *PostMsg )( 
            IApplication * This,
            /* [in] */ LONG msg,
            /* [in] */ LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result);
        
        END_INTERFACE
    } IApplicationVtbl;

    interface IApplication
    {
        CONST_VTBL struct IApplicationVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IApplication_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IApplication_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IApplication_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IApplication_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IApplication_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IApplication_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IApplication_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IApplication_SayHi(This)	\
    (This)->lpVtbl -> SayHi(This)

#define IApplication_get_Playlist(This,pVal)	\
    (This)->lpVtbl -> get_Playlist(This,pVal)

#define IApplication_get_MediaLibrary(This,pVal)	\
    (This)->lpVtbl -> get_MediaLibrary(This,pVal)

#define IApplication_Play(This)	\
    (This)->lpVtbl -> Play(This)

#define IApplication_StopPlayback(This)	\
    (This)->lpVtbl -> StopPlayback(This)

#define IApplication_Pause(This)	\
    (This)->lpVtbl -> Pause(This)

#define IApplication_Previous(This)	\
    (This)->lpVtbl -> Previous(This)

#define IApplication_Skip(This)	\
    (This)->lpVtbl -> Skip(This)

#define IApplication_LoadItem(This,FileName,MediaItem)	\
    (This)->lpVtbl -> LoadItem(This,FileName,MediaItem)

#define IApplication_SetTimeout(This,timeout,timeoutfunction,timerid)	\
    (This)->lpVtbl -> SetTimeout(This,timeout,timeoutfunction,timerid)

#define IApplication_CancelTimer(This,timerid)	\
    (This)->lpVtbl -> CancelTimer(This,timerid)

#define IApplication_GetIniFile(This,inifilename)	\
    (This)->lpVtbl -> GetIniFile(This,inifilename)

#define IApplication_GetIniDirectory(This,iniDirectory)	\
    (This)->lpVtbl -> GetIniDirectory(This,iniDirectory)

#define IApplication_ExecVisPlugin(This,VisDllFile)	\
    (This)->lpVtbl -> ExecVisPlugin(This,VisDllFile)

#define IApplication_get_Skin(This,pVal)	\
    (This)->lpVtbl -> get_Skin(This,pVal)

#define IApplication_put_Skin(This,newVal)	\
    (This)->lpVtbl -> put_Skin(This,newVal)

#define IApplication_get_Shuffle(This,pVal)	\
    (This)->lpVtbl -> get_Shuffle(This,pVal)

#define IApplication_put_Shuffle(This,newVal)	\
    (This)->lpVtbl -> put_Shuffle(This,newVal)

#define IApplication_get_Repeat(This,pVal)	\
    (This)->lpVtbl -> get_Repeat(This,pVal)

#define IApplication_put_Repeat(This,newVal)	\
    (This)->lpVtbl -> put_Repeat(This,newVal)

#define IApplication_RestartWinamp(This)	\
    (This)->lpVtbl -> RestartWinamp(This)

#define IApplication_ShowNotification(This)	\
    (This)->lpVtbl -> ShowNotification(This)

#define IApplication_GetWaVersion(This,version)	\
    (This)->lpVtbl -> GetWaVersion(This,version)

#define IApplication_GetSendToItems(This,Items)	\
    (This)->lpVtbl -> GetSendToItems(This,Items)

#define IApplication_get_Volume(This,pVal)	\
    (This)->lpVtbl -> get_Volume(This,pVal)

#define IApplication_put_Volume(This,newVal)	\
    (This)->lpVtbl -> put_Volume(This,newVal)

#define IApplication_get_Panning(This,pVal)	\
    (This)->lpVtbl -> get_Panning(This,pVal)

#define IApplication_put_Panning(This,newVal)	\
    (This)->lpVtbl -> put_Panning(This,newVal)

#define IApplication_get_PlayState(This,pVal)	\
    (This)->lpVtbl -> get_PlayState(This,pVal)

#define IApplication_get_Position(This,pVal)	\
    (This)->lpVtbl -> get_Position(This,pVal)

#define IApplication_put_Position(This,newVal)	\
    (This)->lpVtbl -> put_Position(This,newVal)

#define IApplication_RunScript(This,scriptfile,arguments)	\
    (This)->lpVtbl -> RunScript(This,scriptfile,arguments)

#define IApplication_UpdateTitle(This)	\
    (This)->lpVtbl -> UpdateTitle(This)

#define IApplication_get_Hwnd(This,pVal)	\
    (This)->lpVtbl -> get_Hwnd(This,pVal)

#define IApplication_SendMsg(This,msg,wParam,lParam,result)	\
    (This)->lpVtbl -> SendMsg(This,msg,wParam,lParam,result)

#define IApplication_PostMsg(This,msg,wParam,lParam,result)	\
    (This)->lpVtbl -> PostMsg(This,msg,wParam,lParam,result)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_SayHi_Proxy( 
    IApplication * This);


void __RPC_STUB IApplication_SayHi_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IApplication_get_Playlist_Proxy( 
    IApplication * This,
    /* [retval][out] */ IPlaylist **pVal);


void __RPC_STUB IApplication_get_Playlist_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IApplication_get_MediaLibrary_Proxy( 
    IApplication * This,
    /* [retval][out] */ IMediaLibrary **pVal);


void __RPC_STUB IApplication_get_MediaLibrary_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_Play_Proxy( 
    IApplication * This);


void __RPC_STUB IApplication_Play_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_StopPlayback_Proxy( 
    IApplication * This);


void __RPC_STUB IApplication_StopPlayback_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_Pause_Proxy( 
    IApplication * This);


void __RPC_STUB IApplication_Pause_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_Previous_Proxy( 
    IApplication * This);


void __RPC_STUB IApplication_Previous_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_Skip_Proxy( 
    IApplication * This);


void __RPC_STUB IApplication_Skip_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_LoadItem_Proxy( 
    IApplication * This,
    /* [in] */ BSTR FileName,
    /* [retval][out] */ IMediaItem **MediaItem);


void __RPC_STUB IApplication_LoadItem_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_SetTimeout_Proxy( 
    IApplication * This,
    /* [in] */ LONG timeout,
    IDispatch *timeoutfunction,
    /* [retval][out] */ LONG *timerid);


void __RPC_STUB IApplication_SetTimeout_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_CancelTimer_Proxy( 
    IApplication * This,
    /* [in] */ LONG timerid);


void __RPC_STUB IApplication_CancelTimer_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_GetIniFile_Proxy( 
    IApplication * This,
    /* [retval][out] */ BSTR *inifilename);


void __RPC_STUB IApplication_GetIniFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_GetIniDirectory_Proxy( 
    IApplication * This,
    /* [retval][out] */ BSTR *iniDirectory);


void __RPC_STUB IApplication_GetIniDirectory_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_ExecVisPlugin_Proxy( 
    IApplication * This,
    /* [in] */ BSTR VisDllFile);


void __RPC_STUB IApplication_ExecVisPlugin_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IApplication_get_Skin_Proxy( 
    IApplication * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IApplication_get_Skin_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IApplication_put_Skin_Proxy( 
    IApplication * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IApplication_put_Skin_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IApplication_get_Shuffle_Proxy( 
    IApplication * This,
    /* [retval][out] */ VARIANT_BOOL *pVal);


void __RPC_STUB IApplication_get_Shuffle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IApplication_put_Shuffle_Proxy( 
    IApplication * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IApplication_put_Shuffle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IApplication_get_Repeat_Proxy( 
    IApplication * This,
    /* [retval][out] */ VARIANT_BOOL *pVal);


void __RPC_STUB IApplication_get_Repeat_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IApplication_put_Repeat_Proxy( 
    IApplication * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IApplication_put_Repeat_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_RestartWinamp_Proxy( 
    IApplication * This);


void __RPC_STUB IApplication_RestartWinamp_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_ShowNotification_Proxy( 
    IApplication * This);


void __RPC_STUB IApplication_ShowNotification_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_GetWaVersion_Proxy( 
    IApplication * This,
    /* [retval][out] */ LONG *version);


void __RPC_STUB IApplication_GetWaVersion_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_GetSendToItems_Proxy( 
    IApplication * This,
    /* [retval][out] */ VARIANT *Items);


void __RPC_STUB IApplication_GetSendToItems_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IApplication_get_Volume_Proxy( 
    IApplication * This,
    /* [retval][out] */ BYTE *pVal);


void __RPC_STUB IApplication_get_Volume_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IApplication_put_Volume_Proxy( 
    IApplication * This,
    /* [in] */ BYTE newVal);


void __RPC_STUB IApplication_put_Volume_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IApplication_get_Panning_Proxy( 
    IApplication * This,
    /* [retval][out] */ INT *pVal);


void __RPC_STUB IApplication_get_Panning_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IApplication_put_Panning_Proxy( 
    IApplication * This,
    /* [in] */ INT newVal);


void __RPC_STUB IApplication_put_Panning_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IApplication_get_PlayState_Proxy( 
    IApplication * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IApplication_get_PlayState_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IApplication_get_Position_Proxy( 
    IApplication * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IApplication_get_Position_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IApplication_put_Position_Proxy( 
    IApplication * This,
    /* [in] */ LONG newVal);


void __RPC_STUB IApplication_put_Position_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_RunScript_Proxy( 
    IApplication * This,
    /* [in] */ BSTR scriptfile,
    /* [in] */ BSTR arguments);


void __RPC_STUB IApplication_RunScript_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_UpdateTitle_Proxy( 
    IApplication * This);


void __RPC_STUB IApplication_UpdateTitle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IApplication_get_Hwnd_Proxy( 
    IApplication * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IApplication_get_Hwnd_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_SendMsg_Proxy( 
    IApplication * This,
    /* [in] */ LONG msg,
    /* [in] */ LONG wParam,
    /* [in] */ LONGLONG lParam,
    /* [retval][out] */ LONG *result);


void __RPC_STUB IApplication_SendMsg_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IApplication_PostMsg_Proxy( 
    IApplication * This,
    /* [in] */ LONG msg,
    /* [in] */ LONG wParam,
    /* [in] */ LONGLONG lParam,
    /* [retval][out] */ LONG *result);


void __RPC_STUB IApplication_PostMsg_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IApplication_INTERFACE_DEFINED__ */


#ifndef __IMediaItem_INTERFACE_DEFINED__
#define __IMediaItem_INTERFACE_DEFINED__

/* interface IMediaItem */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IMediaItem;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("E7B30607-7180-4E40-A5C7-AF9F7D1C30C7")
    IMediaItem : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Filename( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Position( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Position( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Title( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Title( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ATFString( 
            /* [in] */ BSTR ATFSpecification,
            /* [retval][out] */ BSTR *FormattedString) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Enqueue( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Artist( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Artist( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Album( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Album( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Rating( 
            /* [retval][out] */ BYTE *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Rating( 
            /* [in] */ BYTE newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Playcount( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Playcount( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Insert( 
            LONG index) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LastPlay( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_LastPlay( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RefreshMeta( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DbIndex( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Length( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Track( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Genre( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Genre( 
            /* [in] */ BSTR newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IMediaItemVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMediaItem * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMediaItem * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMediaItem * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMediaItem * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMediaItem * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMediaItem * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMediaItem * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Name )( 
            IMediaItem * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Filename )( 
            IMediaItem * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Position )( 
            IMediaItem * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Position )( 
            IMediaItem * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Title )( 
            IMediaItem * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Title )( 
            IMediaItem * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ATFString )( 
            IMediaItem * This,
            /* [in] */ BSTR ATFSpecification,
            /* [retval][out] */ BSTR *FormattedString);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Enqueue )( 
            IMediaItem * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Artist )( 
            IMediaItem * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Artist )( 
            IMediaItem * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Album )( 
            IMediaItem * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Album )( 
            IMediaItem * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Rating )( 
            IMediaItem * This,
            /* [retval][out] */ BYTE *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Rating )( 
            IMediaItem * This,
            /* [in] */ BYTE newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Playcount )( 
            IMediaItem * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Playcount )( 
            IMediaItem * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Insert )( 
            IMediaItem * This,
            LONG index);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LastPlay )( 
            IMediaItem * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_LastPlay )( 
            IMediaItem * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *RefreshMeta )( 
            IMediaItem * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_DbIndex )( 
            IMediaItem * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Length )( 
            IMediaItem * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Track )( 
            IMediaItem * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Genre )( 
            IMediaItem * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Genre )( 
            IMediaItem * This,
            /* [in] */ BSTR newVal);
        
        END_INTERFACE
    } IMediaItemVtbl;

    interface IMediaItem
    {
        CONST_VTBL struct IMediaItemVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMediaItem_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IMediaItem_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IMediaItem_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IMediaItem_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IMediaItem_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IMediaItem_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IMediaItem_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IMediaItem_get_Name(This,pVal)	\
    (This)->lpVtbl -> get_Name(This,pVal)

#define IMediaItem_get_Filename(This,pVal)	\
    (This)->lpVtbl -> get_Filename(This,pVal)

#define IMediaItem_get_Position(This,pVal)	\
    (This)->lpVtbl -> get_Position(This,pVal)

#define IMediaItem_put_Position(This,newVal)	\
    (This)->lpVtbl -> put_Position(This,newVal)

#define IMediaItem_get_Title(This,pVal)	\
    (This)->lpVtbl -> get_Title(This,pVal)

#define IMediaItem_put_Title(This,newVal)	\
    (This)->lpVtbl -> put_Title(This,newVal)

#define IMediaItem_ATFString(This,ATFSpecification,FormattedString)	\
    (This)->lpVtbl -> ATFString(This,ATFSpecification,FormattedString)

#define IMediaItem_Enqueue(This)	\
    (This)->lpVtbl -> Enqueue(This)

#define IMediaItem_get_Artist(This,pVal)	\
    (This)->lpVtbl -> get_Artist(This,pVal)

#define IMediaItem_put_Artist(This,newVal)	\
    (This)->lpVtbl -> put_Artist(This,newVal)

#define IMediaItem_get_Album(This,pVal)	\
    (This)->lpVtbl -> get_Album(This,pVal)

#define IMediaItem_put_Album(This,newVal)	\
    (This)->lpVtbl -> put_Album(This,newVal)

#define IMediaItem_get_Rating(This,pVal)	\
    (This)->lpVtbl -> get_Rating(This,pVal)

#define IMediaItem_put_Rating(This,newVal)	\
    (This)->lpVtbl -> put_Rating(This,newVal)

#define IMediaItem_get_Playcount(This,pVal)	\
    (This)->lpVtbl -> get_Playcount(This,pVal)

#define IMediaItem_put_Playcount(This,newVal)	\
    (This)->lpVtbl -> put_Playcount(This,newVal)

#define IMediaItem_Insert(This,index)	\
    (This)->lpVtbl -> Insert(This,index)

#define IMediaItem_get_LastPlay(This,pVal)	\
    (This)->lpVtbl -> get_LastPlay(This,pVal)

#define IMediaItem_put_LastPlay(This,newVal)	\
    (This)->lpVtbl -> put_LastPlay(This,newVal)

#define IMediaItem_RefreshMeta(This)	\
    (This)->lpVtbl -> RefreshMeta(This)

#define IMediaItem_get_DbIndex(This,pVal)	\
    (This)->lpVtbl -> get_DbIndex(This,pVal)

#define IMediaItem_get_Length(This,pVal)	\
    (This)->lpVtbl -> get_Length(This,pVal)

#define IMediaItem_get_Track(This,pVal)	\
    (This)->lpVtbl -> get_Track(This,pVal)

#define IMediaItem_get_Genre(This,pVal)	\
    (This)->lpVtbl -> get_Genre(This,pVal)

#define IMediaItem_put_Genre(This,newVal)	\
    (This)->lpVtbl -> put_Genre(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Name_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IMediaItem_get_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Filename_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IMediaItem_get_Filename_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Position_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IMediaItem_get_Position_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMediaItem_put_Position_Proxy( 
    IMediaItem * This,
    /* [in] */ LONG newVal);


void __RPC_STUB IMediaItem_put_Position_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Title_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IMediaItem_get_Title_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMediaItem_put_Title_Proxy( 
    IMediaItem * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IMediaItem_put_Title_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMediaItem_ATFString_Proxy( 
    IMediaItem * This,
    /* [in] */ BSTR ATFSpecification,
    /* [retval][out] */ BSTR *FormattedString);


void __RPC_STUB IMediaItem_ATFString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMediaItem_Enqueue_Proxy( 
    IMediaItem * This);


void __RPC_STUB IMediaItem_Enqueue_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Artist_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IMediaItem_get_Artist_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMediaItem_put_Artist_Proxy( 
    IMediaItem * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IMediaItem_put_Artist_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Album_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IMediaItem_get_Album_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMediaItem_put_Album_Proxy( 
    IMediaItem * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IMediaItem_put_Album_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Rating_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ BYTE *pVal);


void __RPC_STUB IMediaItem_get_Rating_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMediaItem_put_Rating_Proxy( 
    IMediaItem * This,
    /* [in] */ BYTE newVal);


void __RPC_STUB IMediaItem_put_Rating_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Playcount_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IMediaItem_get_Playcount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMediaItem_put_Playcount_Proxy( 
    IMediaItem * This,
    /* [in] */ LONG newVal);


void __RPC_STUB IMediaItem_put_Playcount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMediaItem_Insert_Proxy( 
    IMediaItem * This,
    LONG index);


void __RPC_STUB IMediaItem_Insert_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_LastPlay_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IMediaItem_get_LastPlay_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMediaItem_put_LastPlay_Proxy( 
    IMediaItem * This,
    /* [in] */ LONG newVal);


void __RPC_STUB IMediaItem_put_LastPlay_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMediaItem_RefreshMeta_Proxy( 
    IMediaItem * This);


void __RPC_STUB IMediaItem_RefreshMeta_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_DbIndex_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IMediaItem_get_DbIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Length_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IMediaItem_get_Length_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Track_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IMediaItem_get_Track_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaItem_get_Genre_Proxy( 
    IMediaItem * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IMediaItem_get_Genre_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMediaItem_put_Genre_Proxy( 
    IMediaItem * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IMediaItem_put_Genre_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IMediaItem_INTERFACE_DEFINED__ */


#ifndef __IMediaLibrary_INTERFACE_DEFINED__
#define __IMediaLibrary_INTERFACE_DEFINED__

/* interface IMediaLibrary */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IMediaLibrary;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("4033571A-3035-4B64-8A3A-E316691C972C")
    IMediaLibrary : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RunQueryArray( 
            /* [in] */ BSTR QueryString,
            /* [retval][out] */ VARIANT *MediaItems) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetItem( 
            /* [in] */ BSTR Filename,
            /* [retval][out] */ IMediaItem **MediaItem) = 0;
        
        virtual /* [helpstring][id][restricted][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ LONG index,
            /* [retval][out] */ IMediaItem **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Hwnd( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SendMsg( 
            /* [in] */ LONG msg,
            /* [in] */ LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE PostMsg( 
            /* [in] */ LONG msg,
            /* [in] */ LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IMediaLibraryVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMediaLibrary * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMediaLibrary * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMediaLibrary * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMediaLibrary * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMediaLibrary * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMediaLibrary * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMediaLibrary * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *RunQueryArray )( 
            IMediaLibrary * This,
            /* [in] */ BSTR QueryString,
            /* [retval][out] */ VARIANT *MediaItems);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetItem )( 
            IMediaLibrary * This,
            /* [in] */ BSTR Filename,
            /* [retval][out] */ IMediaItem **MediaItem);
        
        /* [helpstring][id][restricted][propget] */ HRESULT ( STDMETHODCALLTYPE *get__NewEnum )( 
            IMediaLibrary * This,
            /* [retval][out] */ IUnknown **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Item )( 
            IMediaLibrary * This,
            /* [in] */ LONG index,
            /* [retval][out] */ IMediaItem **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Hwnd )( 
            IMediaLibrary * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SendMsg )( 
            IMediaLibrary * This,
            /* [in] */ LONG msg,
            /* [in] */ LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *PostMsg )( 
            IMediaLibrary * This,
            /* [in] */ LONG msg,
            /* [in] */ LONG wParam,
            /* [in] */ LONGLONG lParam,
            /* [retval][out] */ LONG *result);
        
        END_INTERFACE
    } IMediaLibraryVtbl;

    interface IMediaLibrary
    {
        CONST_VTBL struct IMediaLibraryVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMediaLibrary_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IMediaLibrary_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IMediaLibrary_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IMediaLibrary_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IMediaLibrary_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IMediaLibrary_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IMediaLibrary_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IMediaLibrary_RunQueryArray(This,QueryString,MediaItems)	\
    (This)->lpVtbl -> RunQueryArray(This,QueryString,MediaItems)

#define IMediaLibrary_GetItem(This,Filename,MediaItem)	\
    (This)->lpVtbl -> GetItem(This,Filename,MediaItem)

#define IMediaLibrary_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define IMediaLibrary_get_Item(This,index,pVal)	\
    (This)->lpVtbl -> get_Item(This,index,pVal)

#define IMediaLibrary_get_Hwnd(This,pVal)	\
    (This)->lpVtbl -> get_Hwnd(This,pVal)

#define IMediaLibrary_SendMsg(This,msg,wParam,lParam,result)	\
    (This)->lpVtbl -> SendMsg(This,msg,wParam,lParam,result)

#define IMediaLibrary_PostMsg(This,msg,wParam,lParam,result)	\
    (This)->lpVtbl -> PostMsg(This,msg,wParam,lParam,result)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMediaLibrary_RunQueryArray_Proxy( 
    IMediaLibrary * This,
    /* [in] */ BSTR QueryString,
    /* [retval][out] */ VARIANT *MediaItems);


void __RPC_STUB IMediaLibrary_RunQueryArray_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMediaLibrary_GetItem_Proxy( 
    IMediaLibrary * This,
    /* [in] */ BSTR Filename,
    /* [retval][out] */ IMediaItem **MediaItem);


void __RPC_STUB IMediaLibrary_GetItem_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][restricted][propget] */ HRESULT STDMETHODCALLTYPE IMediaLibrary_get__NewEnum_Proxy( 
    IMediaLibrary * This,
    /* [retval][out] */ IUnknown **pVal);


void __RPC_STUB IMediaLibrary_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaLibrary_get_Item_Proxy( 
    IMediaLibrary * This,
    /* [in] */ LONG index,
    /* [retval][out] */ IMediaItem **pVal);


void __RPC_STUB IMediaLibrary_get_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMediaLibrary_get_Hwnd_Proxy( 
    IMediaLibrary * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IMediaLibrary_get_Hwnd_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMediaLibrary_SendMsg_Proxy( 
    IMediaLibrary * This,
    /* [in] */ LONG msg,
    /* [in] */ LONG wParam,
    /* [in] */ LONGLONG lParam,
    /* [retval][out] */ LONG *result);


void __RPC_STUB IMediaLibrary_SendMsg_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMediaLibrary_PostMsg_Proxy( 
    IMediaLibrary * This,
    /* [in] */ LONG msg,
    /* [in] */ LONG wParam,
    /* [in] */ LONGLONG lParam,
    /* [retval][out] */ LONG *result);


void __RPC_STUB IMediaLibrary_PostMsg_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IMediaLibrary_INTERFACE_DEFINED__ */


#ifndef __ISiteManager_INTERFACE_DEFINED__
#define __ISiteManager_INTERFACE_DEFINED__

/* interface ISiteManager */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ISiteManager;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("1C95F212-3B28-4713-B6AB-35C422409D75")
    ISiteManager : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AttachEvents( 
            /* [in] */ IDispatch *EventObject,
            BSTR ObjectPrefix) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Quit( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Description( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Description( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Arguments( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ISiteManagerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ISiteManager * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ISiteManager * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ISiteManager * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ISiteManager * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ISiteManager * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ISiteManager * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ISiteManager * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *AttachEvents )( 
            ISiteManager * This,
            /* [in] */ IDispatch *EventObject,
            BSTR ObjectPrefix);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Quit )( 
            ISiteManager * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Description )( 
            ISiteManager * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Description )( 
            ISiteManager * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Arguments )( 
            ISiteManager * This,
            /* [retval][out] */ BSTR *pVal);
        
        END_INTERFACE
    } ISiteManagerVtbl;

    interface ISiteManager
    {
        CONST_VTBL struct ISiteManagerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISiteManager_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ISiteManager_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ISiteManager_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ISiteManager_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ISiteManager_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ISiteManager_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ISiteManager_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ISiteManager_AttachEvents(This,EventObject,ObjectPrefix)	\
    (This)->lpVtbl -> AttachEvents(This,EventObject,ObjectPrefix)

#define ISiteManager_Quit(This)	\
    (This)->lpVtbl -> Quit(This)

#define ISiteManager_get_Description(This,pVal)	\
    (This)->lpVtbl -> get_Description(This,pVal)

#define ISiteManager_put_Description(This,newVal)	\
    (This)->lpVtbl -> put_Description(This,newVal)

#define ISiteManager_get_Arguments(This,pVal)	\
    (This)->lpVtbl -> get_Arguments(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ISiteManager_AttachEvents_Proxy( 
    ISiteManager * This,
    /* [in] */ IDispatch *EventObject,
    BSTR ObjectPrefix);


void __RPC_STUB ISiteManager_AttachEvents_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ISiteManager_Quit_Proxy( 
    ISiteManager * This);


void __RPC_STUB ISiteManager_Quit_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISiteManager_get_Description_Proxy( 
    ISiteManager * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ISiteManager_get_Description_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ISiteManager_put_Description_Proxy( 
    ISiteManager * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB ISiteManager_put_Description_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISiteManager_get_Arguments_Proxy( 
    ISiteManager * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ISiteManager_get_Arguments_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ISiteManager_INTERFACE_DEFINED__ */



#ifndef __ActiveWinamp_LIBRARY_DEFINED__
#define __ActiveWinamp_LIBRARY_DEFINED__

/* library ActiveWinamp */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_ActiveWinamp;

#ifndef ___IApplicationEvents_DISPINTERFACE_DEFINED__
#define ___IApplicationEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IApplicationEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IApplicationEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("A1B0673A-2632-4FCA-8091-4CFD1F2A7710")
    _IApplicationEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IApplicationEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IApplicationEvents * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IApplicationEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IApplicationEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IApplicationEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IApplicationEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IApplicationEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IApplicationEvents * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } _IApplicationEventsVtbl;

    interface _IApplicationEvents
    {
        CONST_VTBL struct _IApplicationEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IApplicationEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _IApplicationEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _IApplicationEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _IApplicationEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _IApplicationEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _IApplicationEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _IApplicationEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IApplicationEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_Application;

#ifdef __cplusplus

class DECLSPEC_UUID("1C7F39AF-65C0-4C14-A392-6B4714705DC2")
Application;
#endif

EXTERN_C const CLSID CLSID_Playlist;

#ifdef __cplusplus

class DECLSPEC_UUID("5108E2F5-A7E8-4284-9555-9FA8C3994B9C")
Playlist;
#endif

EXTERN_C const CLSID CLSID_MediaItem;

#ifdef __cplusplus

class DECLSPEC_UUID("CF07CEDE-5BA9-414D-87C1-28737EDEE17D")
MediaItem;
#endif

EXTERN_C const CLSID CLSID_MediaLibrary;

#ifdef __cplusplus

class DECLSPEC_UUID("9A01F0B7-47F9-41F6-93C4-1B2ED242DA4B")
MediaLibrary;
#endif

EXTERN_C const CLSID CLSID_SiteManager;

#ifdef __cplusplus

class DECLSPEC_UUID("89BBB22F-117F-4548-A2F8-FFDEBDF456D2")
SiteManager;
#endif
#endif /* __ActiveWinamp_LIBRARY_DEFINED__ */

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


