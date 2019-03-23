/************************************************************************

MediaLibrary.cpp

Implementation of MediaLibrary COM object.

TODO:
Skip method of enumerator
Fix param for sendmsg/postmsg?
Batch processing of library items somehow?
Add/remove items from library?

************************************************************************/

#include "stdafx.h"
#include "MediaLibrary.h"
#include ".\medialibrary.h"
#include "MediaItem.h"

HWND HwndMl;
LONG libmlipc;

// CMediaLibrary
STDMETHODIMP CMediaLibrary::RunQueryArray(BSTR QueryString, VARIANT* MediaItems)
{
	USES_CONVERSION;
	char *tmp = NULL;
	mlQueryStruct mlQ;
	SAFEARRAYBOUND mlsafeArrayBounds[1];
	LONG mlsfaindex;

	if (!HwndMl)
		HwndMl = (HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,-1,(LPARAM)libmlipc);

	if (HwndMl)
	{
		mlQ.query=OLE2A(QueryString);
		mlQ.max_results=0;
		mlQ.results.Alloc=0;
		mlQ.results.Items=NULL;
		mlQ.results.Size=0;

		mlsfaindex = 1;
		
		if (SendMessage(HwndMl,WM_ML_IPC,(WPARAM)&mlQ,ML_IPC_DB_RUNQUERY) == 1)
		{
			MediaItems->vt = VT_ARRAY | VT_VARIANT;
			mlsafeArrayBounds[0].lLbound = 1;
			mlsafeArrayBounds[0].cElements  = mlQ.results.Size;
			MediaItems->parray = SafeArrayCreate(VT_VARIANT, 1, mlsafeArrayBounds);

			for (int i = 0; i < mlQ.results.Size; i++)
			{

				VARIANT xv;
				CComObject<CMediaItem> *plp = new CComObject<CMediaItem>;
				if (plp->CreateFromMLItem(&mlQ.results.Items[i]))
				{
					delete plp;
					return S_FALSE;
				}
				CComQIPtr<IDispatch> ptr2(plp);

				xv.vt = VT_DISPATCH;
				xv.pdispVal = ptr2.Detach();
				mlsafeArrayBounds[0].cElements = mlsfaindex;
				HRESULT hr = SafeArrayPutElement(MediaItems->parray, &mlsfaindex, &xv);
				VariantClear(&xv);
				mlsfaindex++;
			}
			SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlQ, ML_IPC_DB_FREEQUERYRESULTS);
		}
	}

	return S_OK;
}

STDMETHODIMP CMediaLibrary::GetItem(BSTR Filename, IMediaItem** MediaItem)
{
	USES_CONVERSION;

	if (!HwndMl)
		HwndMl = (HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,-1,(LPARAM)libmlipc);

	mlQueryStruct mlQ;
	mlQ.query=OLE2A(Filename);
	mlQ.max_results=1;
	mlQ.results.Alloc=0;
	mlQ.results.Items=NULL;
	mlQ.results.Size=0;
	if (SendMessage(HwndMl,WM_ML_IPC,(WPARAM)&mlQ,ML_IPC_DB_RUNQUERY_FILENAME) == 1)
	{
		CComObject<CMediaItem> *plp = new CComObject<CMediaItem>;
		if (plp->CreateFromMLItem(&mlQ.results.Items[0]))
		{
			delete plp;
			return S_FALSE;
		}
		plp->QueryInterface(IID_IDispatch, (void **)MediaItem); 
		return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CMediaLibrary::get__NewEnum(IUnknown** pVal)
{
	//*pVal = this->GetUnknown();
	CComPtr<CEnumMediaLibrary> pEnum = new CComObject<CEnumMediaLibrary>;
	pEnum->QueryInterface(IID_IUnknown, (void **)pVal);
	return S_OK;
}


STDMETHODIMP CEnumMediaLibrary::Next( 
								 unsigned long  celt,              
								 VARIANT FAR*  rgVar,              
								 unsigned long FAR*  pCeltFetched  
								 )
{
	mlQueryStruct mlQ;

	if (!HwndMl)
		HwndMl = (HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,-1,(LPARAM)libmlipc);

	//char idx[10];
	if (pCeltFetched != NULL)
		*pCeltFetched=0;

	//itoa(m_enumidx, idx, 10);
	mlQ.query=(char*)m_enumidx;
	mlQ.max_results=1;
	mlQ.results.Alloc=0;
	mlQ.results.Items=NULL;
	mlQ.results.Size=0;

	if (SendMessage(HwndMl,WM_ML_IPC,(WPARAM)&mlQ,ML_IPC_DB_RUNQUERY_INDEX) == 1)
	{
		if (mlQ.results.Size != 1)
			return S_FALSE;

		CComObject<CMediaItem> *plp = new CComObject<CMediaItem>;
		if (plp->CreateFromMLItem(&mlQ.results.Items[0]))
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
		SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlQ, ML_IPC_DB_FREEQUERYRESULTS);
		return S_OK;
	} else
		return S_FALSE;
	
}

STDMETHODIMP CMediaLibrary::get_Item(LONG index, IMediaItem** pVal)
{
	mlQueryStruct mlQ;

	if (!HwndMl)
		HwndMl = (HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,-1,(LPARAM)libmlipc);

	//itoa(m_enumidx, idx, 10);
	mlQ.query=(char*)index;
	mlQ.max_results=1;
	mlQ.results.Alloc=0;
	mlQ.results.Items=NULL;
	mlQ.results.Size=0;
	if (SendMessage(HwndMl,WM_ML_IPC,(WPARAM)&mlQ,ML_IPC_DB_RUNQUERY_INDEX) == 1)
	{
		if (mlQ.results.Size != 1)
			return S_FALSE;

		CComObject<CMediaItem> *plp = new CComObject<CMediaItem>;
		if (plp->CreateFromMLItem(&mlQ.results.Items[0]))
		{
			delete plp;
			return S_FALSE;
		}
		plp->QueryInterface(IID_IDispatch, (void **)pVal); 
		SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlQ, ML_IPC_DB_FREEQUERYRESULTS);
	}
	return S_OK;
}

STDMETHODIMP CMediaLibrary::get_Hwnd(LONG* pVal)
{
	if (!HwndMl)
		HwndMl = (HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,-1,(LPARAM)libmlipc);

	*pVal = (LONG)HwndMl;
	return S_OK;
}

STDMETHODIMP CMediaLibrary::SendMsg(LONG msg, LONG wParam, LONGLONG lParam, LONG* result)
{
	if (!HwndMl)
		HwndMl = (HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,-1,(LPARAM)libmlipc);

	*result = SendMessage(HwndMl, (UINT)msg, (WPARAM)wParam, (LONG)lParam);
	return S_OK;
}

STDMETHODIMP CMediaLibrary::PostMsg(LONG msg, LONG wParam, LONGLONG lParam, LONG* result)
{
	if (!HwndMl)
		HwndMl = (HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,-1,(LPARAM)libmlipc);

	*result = PostMessage(HwndMl, (UINT)msg, (WPARAM)wParam, (LPARAM)lParam);
	return S_OK;
}
