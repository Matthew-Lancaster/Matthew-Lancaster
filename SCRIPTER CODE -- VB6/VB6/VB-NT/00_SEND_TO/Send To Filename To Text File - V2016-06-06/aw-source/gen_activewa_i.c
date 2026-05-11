

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


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

#if !defined(_M_IA64) && !defined(_M_AMD64)


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

MIDL_DEFINE_GUID(IID, IID_IPlaylist,0x83F2DBE6,0xFA22,0x4D6B,0xA2,0x81,0x04,0x27,0xD7,0xAF,0x49,0x0C);


MIDL_DEFINE_GUID(IID, IID_IApplication,0x2EBD7857,0xB229,0x4247,0x9F,0xAA,0x17,0xC5,0x05,0x41,0x04,0x74);


MIDL_DEFINE_GUID(IID, IID_IMediaItem,0xE7B30607,0x7180,0x4E40,0xA5,0xC7,0xAF,0x9F,0x7D,0x1C,0x30,0xC7);


MIDL_DEFINE_GUID(IID, IID_IMediaLibrary,0x4033571A,0x3035,0x4B64,0x8A,0x3A,0xE3,0x16,0x69,0x1C,0x97,0x2C);


MIDL_DEFINE_GUID(IID, IID_ISiteManager,0x1C95F212,0x3B28,0x4713,0xB6,0xAB,0x35,0xC4,0x22,0x40,0x9D,0x75);


MIDL_DEFINE_GUID(IID, LIBID_ActiveWinamp,0x142FF258,0xEE9C,0x4527,0xB2,0xC7,0x4E,0xAD,0x10,0xB7,0x52,0xD9);


MIDL_DEFINE_GUID(IID, DIID__IApplicationEvents,0xA1B0673A,0x2632,0x4FCA,0x80,0x91,0x4C,0xFD,0x1F,0x2A,0x77,0x10);


MIDL_DEFINE_GUID(CLSID, CLSID_Application,0x1C7F39AF,0x65C0,0x4C14,0xA3,0x92,0x6B,0x47,0x14,0x70,0x5D,0xC2);


MIDL_DEFINE_GUID(CLSID, CLSID_Playlist,0x5108E2F5,0xA7E8,0x4284,0x95,0x55,0x9F,0xA8,0xC3,0x99,0x4B,0x9C);


MIDL_DEFINE_GUID(CLSID, CLSID_MediaItem,0xCF07CEDE,0x5BA9,0x414D,0x87,0xC1,0x28,0x73,0x7E,0xDE,0xE1,0x7D);


MIDL_DEFINE_GUID(CLSID, CLSID_MediaLibrary,0x9A01F0B7,0x47F9,0x41F6,0x93,0xC4,0x1B,0x2E,0xD2,0x42,0xDA,0x4B);


MIDL_DEFINE_GUID(CLSID, CLSID_SiteManager,0x89BBB22F,0x117F,0x4548,0xA2,0xF8,0xFF,0xDE,0xBD,0xF4,0x56,0xD2);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



#endif /* !defined(_M_IA64) && !defined(_M_AMD64)*/

