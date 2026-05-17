/************************************************************************

Playlist.cpp

Implementation of the playlist COM object. (Application.Playlist).

TODO:
Implement "Skip" for enumeration
Fix LONGLONG param for SendMsg/PostMsg calls?

************************************************************************/

#include "stdafx.h"
#include "Playlist.h"
#include ".\playlist.h"
#include "ipc_pe.h"

#include "MediaItem.h"
extern HWND HwndPl;
extern WNDPROC lpMainWndProcOld;
// CPlaylist

STDMETHODIMP CPlaylist::get_Count(LONG* pVal)
{
	*pVal = SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_GETLISTLENGTH);
	return S_OK;
}

STDMETHODIMP CPlaylist::get__NewEnum(IUnknown** pVal)
{
	//*pVal = this->GetUnknown();
	CComPtr<CEnumPlaylist> pEnum = new CComObject<CEnumPlaylist>;
	pEnum->QueryInterface(IID_IUnknown, (void **)pVal);
	return S_OK;
}
STDMETHODIMP CPlaylist::get_Item(LONG index, IMediaItem** pVal)
{
	CComObject<CMediaItem> *plp = new CComObject<CMediaItem>;
	if (plp->CreateFromPlaylist(index-1))
	{
		delete plp;
		return S_FALSE;
	}
	plp->QueryInterface(IID_IDispatch, (void **)pVal); 
	return S_OK;
}


STDMETHODIMP CEnumPlaylist::Next( 
				unsigned long  celt,              
				VARIANT FAR*  rgVar,              
				unsigned long FAR*  pCeltFetched  
				)
{
	//Needs some work to be able to return more than 1 item.

	//This should be using an array of pointers to IMediaItems.
	//They should only be created as needed to save resources
	
	m_maxitems = SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_GETLISTLENGTH);
	long idx=0;
	if (pCeltFetched != NULL)
		*pCeltFetched=0;

	CComObject<CMediaItem> *plp = new CComObject<CMediaItem>;
	if (plp->CreateFromPlaylist(m_enumidx))
	{
		delete plp;
		return S_FALSE;
	}
	CComQIPtr<IDispatch> ptr2(plp);
		
	VariantInit(&rgVar[0]);
	rgVar[0].vt = VT_DISPATCH;
	rgVar[0].pdispVal = ptr2.Detach();

	if (pCeltFetched != NULL)
		*pCeltFetched=1;

	m_enumidx++;
	return S_OK;
}

STDMETHODIMP CPlaylist::get_Position(LONG* pVal)
{
	*pVal = (SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_GETLISTPOS) + 1);
	return S_OK;
}

STDMETHODIMP CPlaylist::put_Position(LONG newVal)
{
	SendMessage(plugin.hwndParent,WM_WA_IPC,(newVal-1),IPC_SETPLAYLISTPOS);
	return S_OK;
}

SAFEARRAYBOUND safeArrayBounds[1];
LONG sfaindex;
VARIANT* pSelectionArray;
bool bHooking;
LRESULT CALLBACK MainWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	if (bHooking && message == WM_USER && lParam == IPC_HOOK_TITLES)
	{
		int secs = 0;
		char *tokval = NULL;
		long retval = 0;

		DWORD plpos = *((DWORD*)wParam + 8);
		waHookTitleStruct *wah = (waHookTitleStruct *)wParam;
		fileinfo2 fi;
		fi.fileindex = plpos;
		fi.filelength[0] = '\0';
		SendMessage(HwndPl, WM_WA_IPC, IPC_PE_GETINDEXTITLE, (LPARAM)&fi);

		if (fi.filelength[0] == '\0')		
			retval = CallWindowProc(lpMainWndProcOld,hwnd,message,wParam,lParam);

		VARIANT xv;
		CComObject<CMediaItem> *plp = new CComObject<CMediaItem>;
		if (plp->CreateFromFilePlaylist(wah->filename, plpos))
		{
			delete plp;
			return 1;
		}
		
		CComQIPtr<IDispatch> ptr2(plp);

		xv.vt = VT_DISPATCH;
		xv.pdispVal = ptr2.Detach();

		safeArrayBounds[0].cElements = sfaindex;
		SafeArrayRedim(pSelectionArray->parray, safeArrayBounds);
		HRESULT hr = SafeArrayPutElement(pSelectionArray->parray, &sfaindex,&xv);
		VariantClear(&xv);

		if (retval == 0)
		{
			tokval = strtok(fi.filelength, ":");
			if (tokval)
				secs = atoi(tokval);

			while (tokval != NULL)
			{
				tokval = strtok(NULL, ":");
				if (tokval)
					secs = secs * 60 + atoi(tokval);
			}
			wah->length = secs;
			strncpy(wah->title, fi.filetitle, MAX_PATH);
		}

		//Cache what we get
		plp->m_length = wah->length;
		strncpy(plp->m_title, wah->title, MAX_PATH);

		sfaindex++;
		return 1;
	}
	return CallWindowProc(lpMainWndProcOld,hwnd,message,wParam,lParam);
}

STDMETHODIMP CPlaylist::GetSelection(VARIANT* SelectionArray)
{
	safeArrayBounds[0].lLbound = 1;
	safeArrayBounds[0].cElements  = 0;

	SelectionArray->parray = SafeArrayCreate(VT_VARIANT, 1, safeArrayBounds);
	SelectionArray->vt = VT_ARRAY | VT_VARIANT;
	pSelectionArray = SelectionArray;

	sfaindex = 1;

	//Mutex
	lpMainWndProcOld = (WNDPROC)SetWindowLong(plugin.hwndParent, GWL_WNDPROC, (LONG)MainWndProc);
	bHooking = true;
	SendMessage(HwndPl,WM_COMMAND,40293,0);
	bHooking = false;
	(WNDPROC)SetWindowLong(plugin.hwndParent, GWL_WNDPROC, (LONG)lpMainWndProcOld);

	return S_OK;
}

STDMETHODIMP CPlaylist::Clear(void)
{
	SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_DELETE);
	return S_OK;
}

STDMETHODIMP CPlaylist::DeleteIndex(LONG index)
{
	SendMessage(HwndPl,WM_WA_IPC,IPC_PE_DELETEINDEX,index-1);
	return S_OK;
}

STDMETHODIMP CPlaylist::SwapIndex(INT from, INT to)
{
	// (lParam & 0xFFFF0000) >> 16 = from, (lParam & 0xFFFF) = to
	unsigned int lpm = from - 1;
	lpm = (lpm << 16) | (to - 1);

	SendMessage(HwndPl,WM_WA_IPC,IPC_PE_SWAPINDEX,lpm);
	return S_OK;
}


STDMETHODIMP CPlaylist::FlushCache(void)
{
	SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_REFRESHPLCACHE);
	return S_OK;
}

STDMETHODIMP CPlaylist::get_Hwnd(LONG* pVal)
{
	*pVal = (LONG)HwndPl;
	return S_OK;
}

STDMETHODIMP CPlaylist::SendMsg(LONG msg, LONG wParam, LONGLONG lParam, LONG* result)
{
	*result = SendMessage(HwndPl, (UINT)msg, (WPARAM)wParam, (LPARAM)lParam);
	return S_OK;
}

STDMETHODIMP CPlaylist::PostMsg(LONG msg, LONG wParam, LONGLONG lParam, LONG *result)
{
	*result = PostMessage(HwndPl, (UINT)msg, (WPARAM)wParam, (LPARAM)lParam);
	return S_OK;
}
