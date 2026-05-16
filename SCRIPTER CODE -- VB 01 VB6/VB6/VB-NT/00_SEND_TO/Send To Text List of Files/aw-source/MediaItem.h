// MediaItem.h : Declaration of the CMediaItem

#pragma once
#include "resource.h"       // main symbols
#include "ml.h"
#include "gen_activewa.h"
#include "ipc_pe.h"
#include <string>

// CMediaItem

class ATL_NO_VTABLE CMediaItem : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CMediaItem, &CLSID_MediaItem>,
	public IDispatchImpl<IMediaItem, &IID_IMediaItem, &LIBID_ActiveWinamp, /*wMajor =*/ 0xffff, /*wMinor =*/ 0xffff>
{
public:
	CMediaItem()
	{
		m_filename[0] = '\0';
		m_title[0] = '\0';
		m_artist[0] = '\0';
		m_album[0] = '\0';
		m_genre[0] = '\0';
		m_comment[0] = '\0';
		m_plposition = -1;
		m_length = -1;
		m_year = -1;
		m_playcount = -1;
		m_rating = -1;
		m_track = -1;
		m_filesize = -1;
		m_dbidx = -1;
		m_filetime = -1;
		m_type = -1;
		m_lastupd = -1;
		m_bmetaloaded = false;
	}

//DECLARE_REGISTRY_RESOURCEID(IDR_MEDIAITEM)

DECLARE_NOT_AGGREGATABLE(CMediaItem)

BEGIN_COM_MAP(CMediaItem)
	COM_INTERFACE_ENTRY(IMediaItem)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:

	int CreateFromPlaylist(long index);
	int CreateFromMLItem(itemRecord* iteminfo);
	int CreateFromFile(char* filename);
	int CreateFromFilePlaylist(char* filename, long index);

	//Should be std::strings.. yech.
	char m_filename[MAX_PATH];
	char m_title[MAX_PATH];
	char m_album[MAX_PATH];
	char m_artist[MAX_PATH];
	char m_genre[MAX_PATH];
	char m_comment[MAX_PATH];
	int m_plposition;
	int m_length;
	int m_track;
	int m_year;
	int m_rating;
	int m_playcount;
	long m_lastplay;
	int m_bitrate;
	int m_filesize;
	int m_dbidx;
	int m_filetime;
	int m_type;
	long m_lastupd;


	bool m_bmetaloaded;

	STDMETHOD(get_Name)(BSTR* pVal);
	STDMETHOD(get_Filename)(BSTR* pVal);
	STDMETHOD(get_Position)(LONG *pVal);
	STDMETHOD(put_Position)(LONG newVal);
	STDMETHOD(get_Title)(BSTR* pVal);
	STDMETHOD(put_Title)(BSTR newVal);
	STDMETHOD(ATFString)(BSTR ATFSpecification, BSTR* FormattedString);
	
	STDMETHOD(Enqueue)(void);
	
	STDMETHOD(get_Artist)(BSTR* pVal);
	STDMETHOD(put_Artist)(BSTR newVal);
	STDMETHOD(get_Album)(BSTR* pVal);
	STDMETHOD(put_Album)(BSTR newVal);
	STDMETHOD(get_Rating)(BYTE* pVal);
	STDMETHOD(put_Rating)(BYTE newVal);
	STDMETHOD(get_Playcount)(LONG* pVal);
	STDMETHOD(put_Playcount)(LONG newVal);
	STDMETHOD(Insert)(LONG index);
	STDMETHOD(get_LastPlay)(LONG* pVal);
	STDMETHOD(put_LastPlay)(LONG newVal);
	STDMETHOD(RefreshMeta)(void);
	STDMETHOD(get_DbIndex)(LONG* pVal);
	STDMETHOD(get_Length)(LONG* pVal);
	STDMETHOD(get_Track)(LONG* pVal);
	STDMETHOD(get_Genre)(BSTR* pVal);
	STDMETHOD(put_Genre)(BSTR newVal);
};

//OBJECT_ENTRY_AUTO(__uuidof(MediaItem), CMediaItem)
