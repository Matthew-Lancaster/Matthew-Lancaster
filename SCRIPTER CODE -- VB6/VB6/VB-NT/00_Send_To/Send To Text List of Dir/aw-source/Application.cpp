/************************************************************************
Application.cpp

Implementation of the main Application COM object

TODO:
Support multiple instances. see application.h

************************************************************************/ 

#include "stdafx.h"
#include "application.h"
#include ".\application.h"
#include "ScriptManager.h"
#include <map>

extern CScriptManager glbScriptManager;
typedef CComPtr<IDispatch> timp;
typedef std::pair <UINT_PTR, timp> Timer_Pair;
std::map <UINT_PTR, timp> timermap;

STDMETHODIMP CApplication::SayHi(void)
{
	//Maybe change to a version call.
	SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_SHOW_NOTIFICATION);
	return S_OK;
}

STDMETHODIMP CApplication::Play(void)
{
	SendMessage(plugin.hwndParent,WM_COMMAND,40045,0);
	return S_OK;
}

STDMETHODIMP CApplication::get_Playlist(IPlaylist ** pVal)
{
	m_playlist.AddRef();
	m_playlist.QueryInterface(IID_IDispatch, (void **)pVal); 
	return S_OK;
}

STDMETHODIMP CApplication::get_MediaLibrary(IMediaLibrary** pVal)
{
	m_medialibrary.AddRef();
	m_medialibrary.QueryInterface(IID_IDispatch, (void **)pVal); 
	return S_OK;
}

STDMETHODIMP CApplication::StopPlayback(void)
{
	SendMessage(plugin.hwndParent,WM_COMMAND,40047,0);
	return S_OK;
}

STDMETHODIMP CApplication::Pause(void)
{
	SendMessage(plugin.hwndParent,WM_COMMAND,40046,0);
	return S_OK;
}

STDMETHODIMP CApplication::Previous(void)
{
	SendMessage(plugin.hwndParent,WM_COMMAND,40044,0);
	return S_OK;
}

STDMETHODIMP CApplication::LoadItem(BSTR FileName, IMediaItem** MediaItem)
{
	USES_CONVERSION;
	CComObject<CMediaItem> *plp = new CComObject<CMediaItem>;
	if (plp->CreateFromFile(OLE2T(FileName)))
	{
		delete plp;
		return S_FALSE;
	}
	plp->QueryInterface(IID_IDispatch, (void **)MediaItem); 
	return S_OK;
}

STDMETHODIMP CApplication::Skip(void)
{
	SendMessage(plugin.hwndParent,WM_COMMAND,40048,0);
	return S_OK;
}

VOID CALLBACK TimeoutProc(HWND hwnd,
						UINT uMsg,
						UINT_PTR idEvent,
						DWORD dwTime
						)
{
	std::map <UINT_PTR, timp> :: iterator pIter;
	KillTimer(NULL, idEvent);
	pIter = timermap.find(idEvent);

	if ( pIter == timermap.end( ) )
		;
	else
	{
		ATL::CComDispatchDriver(pIter->second).Invoke0( DISPID( DISPID_VALUE ) );
		timermap.erase(pIter);
	}
}


STDMETHODIMP CApplication::SetTimeout(LONG timeout, IDispatch* timeoutfunction,LONG* timerid)
{
	UINT_PTR id = SetTimer(NULL, 0, (UINT)timeout, TimeoutProc);
	*timerid = (LONG)id;
	timermap.insert ( Timer_Pair ( id, timeoutfunction ) );
	return S_OK;
}
STDMETHODIMP CApplication::CancelTimer(LONG timerid)
{
	std::map <UINT_PTR, timp> :: iterator pIter;
	KillTimer(NULL, timerid);
	pIter = timermap.find(timerid);

	if ( pIter == timermap.end( ) )
		;
	else
		timermap.erase(pIter);
	return S_OK;
}

STDMETHODIMP CApplication::GetIniFile(BSTR* inifilename)
{
	USES_CONVERSION;
	*inifilename = SysAllocString(A2OLE((char*)SendMessage(plugin.hwndParent, WM_WA_IPC, 0, IPC_GETINIFILE)));
	return S_OK;
}

STDMETHODIMP CApplication::GetIniDirectory(BSTR* iniDirectory)
{
	USES_CONVERSION;
	*iniDirectory = SysAllocString(A2OLE((char*)SendMessage(plugin.hwndParent, WM_WA_IPC, 0, IPC_GETINIDIRECTORY)));
	return S_OK;
}

STDMETHODIMP CApplication::ExecVisPlugin(BSTR VisDllFile)
{
	USES_CONVERSION;	
	SendMessage(plugin.hwndParent, WM_WA_IPC,(WPARAM)OLE2A(VisDllFile), IPC_EXECPLUG);
	return S_OK;
}

STDMETHODIMP CApplication::get_Skin(BSTR* pVal)
{
	USES_CONVERSION;
	char skinbuff[MAX_PATH];
	SendMessage(plugin.hwndParent, WM_WA_IPC, (WPARAM)skinbuff, IPC_GETSKIN);
	*pVal = SysAllocString(A2OLE(skinbuff));
	return S_OK;
}

STDMETHODIMP CApplication::put_Skin(BSTR newVal)
{
	USES_CONVERSION;
	SendMessage(plugin.hwndParent, WM_WA_IPC,(WPARAM) OLE2A(newVal), IPC_GETSKIN);
	return S_OK;
}

STDMETHODIMP CApplication::get_Shuffle(VARIANT_BOOL* pVal)
{
	*pVal =	SendMessage(plugin.hwndParent, WM_WA_IPC,0,IPC_GET_SHUFFLE);
	return S_OK;
}

STDMETHODIMP CApplication::put_Shuffle(VARIANT_BOOL newVal)
{
	SendMessage(plugin.hwndParent, WM_WA_IPC,newVal,IPC_SET_SHUFFLE);
	return S_OK;
}

STDMETHODIMP CApplication::get_Repeat(VARIANT_BOOL* pVal)
{
	*pVal =	SendMessage(plugin.hwndParent, WM_WA_IPC,0,IPC_GET_REPEAT);
	return S_OK;
}

STDMETHODIMP CApplication::put_Repeat(VARIANT_BOOL newVal)
{
	SendMessage(plugin.hwndParent, WM_WA_IPC,newVal,IPC_SET_REPEAT);
	return S_OK;
}

STDMETHODIMP CApplication::RestartWinamp(void)
{
	SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_RESTARTWINAMP);
	return S_OK;
}

STDMETHODIMP CApplication::ShowNotification(void)
{
	SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_SHOW_NOTIFICATION);
	return S_OK;
}

STDMETHODIMP CApplication::GetWaVersion(LONG* version)
{
	*version = (LONG)SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_GETVERSION);
	return S_OK;
}

STDMETHODIMP CApplication::GetSendToItems(VARIANT* Items)
{
	//*Items = pSendToArray;
	if (pSendToArray.parray)
	{
		Items->vt = VT_ARRAY | VT_VARIANT;
		SafeArrayCopy((SAFEARRAY*)(pSendToArray.parray), (SAFEARRAY**)&Items->parray);
		return S_OK;
	}
	else
	{
		return E_FAIL;
	}
}

int CApplication::CreateSendToItems(int sourcetype, void* data)
{
	if (pSendToArray.parray)
		SafeArrayDestroy(pSendToArray.parray);

	safeArrayBounds[0].lLbound = 1;
	safeArrayBounds[0].cElements  = 0;

	pSendToArray.parray = SafeArrayCreate(VT_VARIANT, 1, safeArrayBounds);
	pSendToArray.vt = VT_ARRAY | VT_VARIANT;

	sfaindex = 1;
	
    VARIANT xv;
	
	if (sourcetype == ML_TYPE_ITEMRECORDLIST)
	{
		itemRecordList *prl = (itemRecordList *)data;
		for (int i = 0; i < prl->Size; i++)
		{

			CComObject<CMediaItem> *plp = new CComObject<CMediaItem>;
			if (plp->CreateFromMLItem(&prl->Items[i]))
			{
				delete plp;
				return 1;
			}
			CComQIPtr<IDispatch> ptr2(plp);

			xv.vt = VT_DISPATCH;
			xv.pdispVal = ptr2.Detach();

			safeArrayBounds[0].cElements = sfaindex;
			SafeArrayRedim(pSendToArray.parray, safeArrayBounds);
			HRESULT hr = SafeArrayPutElement(pSendToArray.parray, &sfaindex,&xv);
			VariantClear(&xv);

			sfaindex++;
		}
	} else if (sourcetype == ML_TYPE_FILENAMES)
	{
		char *nextitem = NULL;
		nextitem = (char*)data;

		while (nextitem[0] != '\0')
		{
			CComObject<CMediaItem> *plp = new CComObject<CMediaItem>;
			if (plp->CreateFromFile(nextitem))
			{
				delete plp;
				return 1;
			}
			CComQIPtr<IDispatch> ptr2(plp);

			xv.vt = VT_DISPATCH;
			xv.pdispVal = ptr2.Detach();

			safeArrayBounds[0].cElements = sfaindex;
			SafeArrayRedim(pSendToArray.parray, safeArrayBounds);
			HRESULT hr = SafeArrayPutElement(pSendToArray.parray, &sfaindex,&xv);
			VariantClear(&xv);
			
			nextitem += strlen(nextitem) + 1;
			sfaindex++;
		}
	}
	return 0;
}

STDMETHODIMP CApplication::get_Volume(BYTE* pVal)
{
	*pVal = (BYTE)SendMessage(plugin.hwndParent, WM_WA_IPC, -666, IPC_SETVOLUME);
	return S_OK;
}

STDMETHODIMP CApplication::put_Volume(BYTE newVal)
{
	SendMessage(plugin.hwndParent, WM_WA_IPC, (WPARAM)newVal, IPC_SETVOLUME);
	return S_OK;
}

STDMETHODIMP CApplication::get_Hwnd(LONG* pVal)
{
	*pVal = (LONG)plugin.hwndParent;
	return S_OK;
}

STDMETHODIMP CApplication::get_PlayState(LONG* pVal)
{
	*pVal = (LONG)SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_ISPLAYING);
	return S_OK;
}

STDMETHODIMP CApplication::SendMsg(LONG msg, LONG wParam, LONGLONG lParam, LONG* result)
{
	*result = SendMessage(plugin.hwndParent, (UINT)msg, (WPARAM)wParam, (LPARAM)lParam);
	return S_OK;
}

STDMETHODIMP CApplication::RunScript(BSTR scriptfile, BSTR arguments)
{
	USES_CONVERSION;
	glbScriptManager.runfile(OLE2T(scriptfile), false, OLE2T(arguments));
	return S_OK;
}

STDMETHODIMP CApplication::get_Panning(INT* pVal)
{
	*pVal = (BYTE)SendMessage(plugin.hwndParent,WM_WA_IPC,-666,IPC_SETPANNING);
	return S_OK;
}

STDMETHODIMP CApplication::put_Panning(INT newVal)
{
	SendMessage(plugin.hwndParent,WM_WA_IPC,newVal,IPC_SETPANNING);
	return S_OK;
}

STDMETHODIMP CApplication::UpdateTitle(void)
{
	SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_UPDTITLE);
	return S_OK;
}

STDMETHODIMP CApplication::PostMsg(LONG msg, LONG wParam, LONGLONG lParam, LONG* result)
{
	*result = PostMessage(plugin.hwndParent, (UINT)msg, (WPARAM)wParam, (LPARAM)lParam);
	return S_OK;
}

STDMETHODIMP CApplication::get_Position(LONG* pVal)
{
	*pVal = (int)SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_GETOUTPUTTIME);
	return S_OK;
}

STDMETHODIMP CApplication::put_Position(LONG newVal)
{
	SendMessage(plugin.hwndParent,WM_WA_IPC,newVal,IPC_JUMPTOTIME);
	return S_OK;
}
