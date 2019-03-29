// MediaLibrary.h : Declaration of the CMediaLibrary

#pragma once
#include "resource.h"       // main symbols

#include "gen_activewa.h"
#include "wa_ipc.h"
#include "gen.h"
#include "ml.h"

extern winampGeneralPurposePlugin plugin;
extern LONG libmlipc;
extern HWND HwndMl;


class FAR CEnumMediaLibrary :
	public CComObjectRootEx<CComSingleThreadModel>,
	public IEnumVARIANT
{
public:

	BEGIN_COM_MAP(CEnumMediaLibrary)
		COM_INTERFACE_ENTRY(IEnumVARIANT)
	END_COM_MAP()

	//static HRESULT Create(SAFEARRAY FAR*, ULONG, CEnumVariant FAR* FAR*); // Creates and intializes Enumerator
	CEnumMediaLibrary()
	{
		m_enumidx=0;
	}
	
	~CEnumMediaLibrary()
	{
	}
	
	STDMETHOD(Clone)(IEnumVARIANT FAR* FAR*  ppEnum ){
		return S_OK;
	}

	STDMETHOD(Reset)()
	{
		m_enumidx=0;
		return S_OK;
	}

	STDMETHOD(Skip)(
		unsigned long  celt  
		)
	{
		m_enumidx+=celt;
		return S_OK;
	}

	STDMETHOD(Next)( 
		unsigned long  celt,              
		VARIANT FAR*  rgVar,              
		unsigned long FAR*  pCeltFetched  
		);

private:
	int m_enumidx;
	int m_maxitems;
}; 


class ATL_NO_VTABLE CMediaLibrary : 
	public CComObjectRootEx<CComSingleThreadModel>,
	//public CComCoClass<CMediaLibrary, &CLSID_MediaLibrary>,
	public IDispatchImpl<IMediaLibrary, &IID_IMediaLibrary, &LIBID_ActiveWinamp, /*wMajor =*/ 0xffff, /*wMinor =*/ 0xffff>
{
public:
	CMediaLibrary()
	{
	}

DECLARE_NOT_AGGREGATABLE(CMediaLibrary)

BEGIN_COM_MAP(CMediaLibrary)
	COM_INTERFACE_ENTRY(IMediaLibrary)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
		//int y = 0;
		//y++; //junk code to place a breakpoint on.
	}

public:

	STDMETHOD(RunQueryArray)(BSTR QueryString, VARIANT* MediaItems);
	STDMETHOD(GetItem)(BSTR Filename, IMediaItem** MediaItem);
	STDMETHOD(get__NewEnum)(IUnknown** pVal);
	STDMETHOD(get_Item)(LONG index, IMediaItem** pVal);
	STDMETHOD(get_Hwnd)(LONG* pVal);
	STDMETHOD(SendMsg)(LONG msg, LONG wParam, LONGLONG lParam, LONG* result);
	STDMETHOD(PostMsg)(LONG msg, LONG wParam, LONGLONG lParam, LONG* result);
};
