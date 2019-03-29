// Playlist.h : Declaration of the CPlaylist

#pragma once
#include "resource.h"       // main symbols

#include "gen_activewa.h"
#include "wa_ipc.h"
#include "gen.h"

extern winampGeneralPurposePlugin plugin;
extern HWND HwndMl;
extern LONG libmlipc;


class FAR CEnumPlaylist :
	public CComObjectRootEx<CComSingleThreadModel>,
	public IEnumVARIANT
{
public:

	BEGIN_COM_MAP(CEnumPlaylist)
		COM_INTERFACE_ENTRY(IEnumVARIANT)
	END_COM_MAP()

	//static HRESULT Create(SAFEARRAY FAR*, ULONG, CEnumVariant FAR* FAR*); // Creates and intializes Enumerator
	CEnumPlaylist()
	{
		m_enumidx=0;
		m_maxitems = (long)SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_GETLISTLENGTH);
	}
	
	~CEnumPlaylist()
	{
	}
	
	STDMETHOD(Clone)(IEnumVARIANT FAR* FAR*  ppEnum ){
		return S_OK;
	}

	STDMETHOD(Reset)()
	{
		m_enumidx=0;
		m_maxitems = (long)SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_GETLISTLENGTH);
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
	long m_enumidx;
	long m_maxitems;
}; 

// CPlaylist

class ATL_NO_VTABLE CPlaylist : 
	public CComObjectRootEx<CComSingleThreadModel>,
	//public CComCoClass<CPlaylist, &CLSID_Playlist>,
	public IDispatchImpl<IPlaylist, &IID_IPlaylist, &LIBID_ActiveWinamp, /*wMajor =*/ 0xffff, /*wMinor =*/ 0xffff>
	//public IEnumVARIANT
{
public:
	CPlaylist()
	{
	}

//DECLARE_REGISTRY_RESOURCEID(IDR_PLAYLIST)


BEGIN_COM_MAP(CPlaylist)
	COM_INTERFACE_ENTRY(IPlaylist)
	COM_INTERFACE_ENTRY(IDispatch)
	//COM_INTERFACE_ENTRY(IEnumVARIANT)
END_COM_MAP()
 

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
		int y = 0;
		 y++;
	}
	
public:

	STDMETHOD(get_Count)(LONG *pVal);
	STDMETHOD(get__NewEnum)(IUnknown** pVal);
	STDMETHOD(get_Item)(LONG index, IMediaItem** pVal);
	STDMETHOD(get_Position)(LONG* pVal);
	STDMETHOD(put_Position)(LONG newVal);
	STDMETHOD(GetSelection)(VARIANT* SelectionArray);
	STDMETHOD(Clear)(void);
	STDMETHOD(DeleteIndex)(LONG index);
	STDMETHOD(SwapIndex)(INT from, INT to);
	STDMETHOD(FlushCache)(void);
	STDMETHOD(get_Hwnd)(LONG* pVal);
	STDMETHOD(SendMsg)(LONG msg, LONG wParam, LONGLONG lParam, LONG* result);
	STDMETHOD(PostMsg)(LONG msg, LONG wParam, LONGLONG lParam, LONG* result);
};

//OBJECT_ENTRY_AUTO(__uuidof(Playlist), CPlaylist)
