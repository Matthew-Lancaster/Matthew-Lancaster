// waobj.h : Declaration of the CApplication

#pragma once
#include "resource.h"       // main symbols

#include "gen_activewa.h"
#include "_IApplicationEvents_CP.h"
#include "gen.h"
#include "Playlist.h"
#include "MediaLibrary.h"
#include "MediaItem.h"

extern LONG libmlipc;
extern HWND HwndMl;
extern winampGeneralPurposePlugin plugin;
extern HWND HwndPl;
//extern WNDPROC lpMainWndProcOld;

// CApplication
class ATL_NO_VTABLE CApplication : 
	public CComObjectRootEx<CComSingleThreadModel>,
	//public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CApplication, &CLSID_Application>,
	public IConnectionPointContainerImpl<CApplication>,
	public CProxy_IApplicationEvents<CApplication>, 
	public IDispatchImpl<IApplication, &IID_IApplication, &LIBID_ActiveWinamp, /*wMajor =*/ 0xffff, /*wMinor =*/ 0xffff>,
	public IProvideClassInfo2Impl<&CLSID_Application, NULL,&LIBID_ActiveWinamp>
{
public:
	CApplication()
	{
		pSendToArray.parray = NULL;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_APPLICATION)
DECLARE_NOT_AGGREGATABLE(CApplication)

BEGIN_COM_MAP(CApplication)
	COM_INTERFACE_ENTRY(IApplication)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
	COM_INTERFACE_ENTRY(IProvideClassInfo)
	COM_INTERFACE_ENTRY(IProvideClassInfo2)
END_COM_MAP()

BEGIN_CONNECTION_POINT_MAP(CApplication)
	CONNECTION_POINT_ENTRY(__uuidof(_IApplicationEvents))
END_CONNECTION_POINT_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_CLASSFACTORY_SINGLETON(CApplication)
	//DECLARE_CLASSFACTORY()

	DWORD dwRegister;
	HRESULT FinalConstruct()
	{
		/*
		Other monikers for multiple instances... broken
		CComPtr<IBindCtx> bind;
		CComPtr<IRunningObjectTable> rot;
		CComPtr<IMoniker> mk;
		HRESULT hr;
		if( SUCCEEDED(hr = ::CreateBindCtx(0, &bind.p)) &&
			SUCCEEDED(hr = bind->GetRunningObjectTable(&rot.p)) &&
			SUCCEEDED(hr = ::CreateItemMoniker(L"!",
			//SUCCEEDED(hr = ::CreateItemMoniker(L"!",
			CComBSTR(CLSID_Application), &mk.p)) )
			//CComBSTR(L"Winamp v1.x"), &mk.p)) )
		{
			// make visible to IIS, but use weak link
			hr = rot->Register(ROTFLAGS_ALLOWANYCLIENT,
				GetUnknown(), mk, &dwRegister);
		}

		what to do...?
*/
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:
	CComObject<CPlaylist> m_playlist;
	CComObject<CMediaLibrary> m_medialibrary;
	
	VARIANT pSendToArray;
	SAFEARRAYBOUND safeArrayBounds[1];
	LONG sfaindex;

	int CreateSendToItems(int sourcetype, void* data);

	STDMETHOD(SayHi)(void);
	STDMETHOD(Play)(void);
	STDMETHOD(get_Playlist)(IPlaylist ** pVal);
	STDMETHOD(get_MediaLibrary)(IMediaLibrary** pVal);
	STDMETHOD(StopPlayback)(void);
	STDMETHOD(Pause)(void);
	STDMETHOD(Previous)(void);
	STDMETHOD(LoadItem)(BSTR FileName, IMediaItem** MediaItem);
	STDMETHOD(Skip)(void);
	STDMETHOD(SetTimeout)(LONG timeout, IDispatch* timeoutfunction, LONG* timerid);
	STDMETHOD(CancelTimer)(LONG timerid);
	STDMETHOD(GetIniFile)(BSTR* inifilename);
	STDMETHOD(GetIniDirectory)(BSTR* iniDirectory);
	STDMETHOD(ExecVisPlugin)(BSTR VisDllFile);
	STDMETHOD(get_Skin)(BSTR* pVal);
	STDMETHOD(put_Skin)(BSTR newVal);
	STDMETHOD(get_Shuffle)(VARIANT_BOOL* pVal);
	STDMETHOD(put_Shuffle)(VARIANT_BOOL newVal);
	STDMETHOD(get_Repeat)(VARIANT_BOOL* pVal);
	STDMETHOD(put_Repeat)(VARIANT_BOOL newVal);
	STDMETHOD(RestartWinamp)(void);
	STDMETHOD(ShowNotification)(void);
	STDMETHOD(GetWaVersion)(LONG* version);
	STDMETHOD(GetSendToItems)(VARIANT* Items);
	STDMETHOD(get_Volume)(BYTE* pVal);
	STDMETHOD(put_Volume)(BYTE newVal);
	STDMETHOD(get_Hwnd)(LONG* pVal);
	STDMETHOD(get_PlayState)(LONG* pVal);
	STDMETHOD(SendMsg)(LONG msg, LONG wParam, LONGLONG lParam, LONG* result);
	STDMETHOD(RunScript)(BSTR scriptfile, BSTR arguments);
	STDMETHOD(get_Panning)(INT* pVal);
	STDMETHOD(put_Panning)(INT newVal);
	STDMETHOD(UpdateTitle)(void);
	STDMETHOD(PostMsg)(LONG msg, LONG wParam, LONGLONG lParam, LONG* result);
	STDMETHOD(get_Position)(LONG* pVal);
	STDMETHOD(put_Position)(LONG newVal);
};

OBJECT_ENTRY_AUTO(__uuidof(Application), CApplication)
