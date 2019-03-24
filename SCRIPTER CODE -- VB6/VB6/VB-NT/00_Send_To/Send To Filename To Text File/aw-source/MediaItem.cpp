/************************************************************************

MediaItem.cpp

Implementation of the MediaItem COM object.

Notes: This object refrains from obtaining meta data until it is requested,
and when it is, its cached. If obtaining meta data requires looking up the ML,
all meta data obtained is cached at that point.

TODO:
Fixed sized buffers for meta data should be std::strings :P
Methods to add meta data to the ML rather than just update iff it exists?
Verify that all codepaths first check the cache before retrieving from the ML or file

************************************************************************/

#include "stdafx.h"
#include "MediaItem.h"
#include ".\mediaitem.h"
#include "gen.h"
#include "wa_ipc.h"
#include "ml.h"

// CMediaItem
extern winampGeneralPurposePlugin plugin;
extern HWND HwndMl;
extern HWND HwndPl;
extern LONG libmlipc;

long lastSong;
char* myTagFunc(char * tag, void * p);
int LoadMeta(CMediaItem* mediai);

char rateStars[] = "*****";

// Chtmlview
void myTagFreeFunc(char * tag, void * p)
{
	GlobalFree((HGLOBAL) tag);
}

char *GetFormattedTitleFromWinamp(char *fmtspec, CMediaItem* curitem)
{
	char *temp;
	waFormatTitle fmt_title;
	char s[MAX_PATH];
	ZeroMemory(&fmt_title, sizeof(waFormatTitle));
	fmt_title.spec = fmtspec;
	fmt_title.TAGFUNC = myTagFunc;
	fmt_title.TAGFREEFUNC = myTagFreeFunc;
	fmt_title.out = s;
	fmt_title.out_len = sizeof(s)-1;
	fmt_title.p = curitem;
	SendMessage(plugin.hwndParent, WM_WA_IPC, (WPARAM)&fmt_title, IPC_FORMAT_TITLE);
	temp = (char*)GlobalAlloc(GPTR, MAX_PATH);
	lstrcpyn(temp, s, MAX_PATH);
	return temp;
}
char *getRecordExtendedItem(itemRecord *item, char *name)
{
	int x=0;
	if (item->extended_info) for (x = 0; item->extended_info[x]; x++)
	{
		if (!lstrcmp(item->extended_info[x],name))
			return item->extended_info[x]+lstrlen(name)+1;
	}
	return NULL;
}
void setRecordExtendedItem(itemRecord *item, char *name, char *value)
{
	int x=0;
	if (item->extended_info) for (x = 0; item->extended_info[x]; x++)
	{
		if (!lstrcmp(item->extended_info[x],name))
		{
			if (lstrlen(value)>lstrlen(item->extended_info[x])+lstrlen(name)+1)
			{
				free(item->extended_info[x]);
				item->extended_info[x]=(char*)malloc(lstrlen(name)+lstrlen(value)+2);
			}
			lstrcpy(item->extended_info[x],name);
			lstrcpy(item->extended_info[x]+lstrlen(name)+1,value);
			return;
		}
	}
	// x=number of valid items.
	item->extended_info=(char**)realloc(item->extended_info,sizeof(char*) * (x+2));
	if (item->extended_info)
	{
		item->extended_info[x]=(char*)malloc(lstrlen(name)+lstrlen(value)+2);
		lstrcpy(item->extended_info[x],name);
		lstrcpy(item->extended_info[x]+lstrlen(name)+1,value);

		item->extended_info[x+1]=0;
	}
}


int getItemVal(char *filename, char *meta)
{
	int pc = -1;
	if (HwndMl)
	{
		mlQueryStruct mlQ;
		mlQ.query=filename;
		mlQ.max_results=1;
		mlQ.results.Alloc=0;
		mlQ.results.Items=NULL;
		mlQ.results.Size=0;
		if (SendMessage(HwndMl,WM_ML_IPC,(WPARAM)&mlQ,ML_IPC_DB_RUNQUERY_FILENAME) == 1)
		{
			char *pcS = getRecordExtendedItem(&mlQ.results.Items[0], meta);

			if (pcS != NULL)
				pc = atoi(pcS); //Requires funny use of libc - modified libctiny handles it ok
		}
		SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlQ, ML_IPC_DB_FREEQUERYRESULTS);
	} else
	{
		//Keep trying to get it
		HwndMl = (HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,0,(LPARAM)libmlipc);
	}
	return pc;
}

void setItemVal(char *filename, char *meta, int newVal)
{
	char numBuf[5];
	if (filename != NULL)
	{
		wsprintf(numBuf, "%d", newVal);
		itemRecord mlItem;
		mlItem.filename=filename;
		mlItem.album=NULL;
		mlItem.artist=NULL;
		mlItem.genre=NULL;
		mlItem.comment=NULL;
		mlItem.title=NULL;
		mlItem.extended_info=NULL;
		mlItem.length=-1;
		mlItem.track=-1;
		mlItem.year=-1;
		setRecordExtendedItem(&mlItem,meta, numBuf);
		SendMessage(HwndMl,WM_ML_IPC,(WPARAM)&mlItem, ML_IPC_DB_UPDATEITEM);
	}
}

char* myTagFunc(char * tag, void * p)
{
	char *s;
	char retbuf[MAX_PATH];
	ZeroMemory(retbuf, MAX_PATH);
	extendedFileInfoStruct efi;
	efi.retlen = sizeof(retbuf);
	efi.ret = retbuf;

	mlQueryStruct mlQ;
	int nm;
	int nm2;
	unsigned int tp;
	
	CMediaItem* curitem = (CMediaItem*)p;
	efi.filename = curitem->m_filename;

	long tmpval;
	char *tmpalloc;
	char *tmpalloc2;
	char *ls;
	char *rs;
	char nmbuf[10];
	char op[2];

	if (!lstrcmpi(tag, "FILENAME")) {
		lstrcpyn(retbuf, curitem->m_filename, MAX_PATH);
	} else if (!lstrcmpi(tag, "pllen")) {
		wsprintf(retbuf, "%d", SendMessage((HWND)p, WM_WA_IPC, 0,IPC_GETLISTLENGTH));
	} else if (!lstrcmpi(tag, "plpos")) {
		nm = SendMessage((HWND)p, WM_WA_IPC, 0,IPC_GETLISTPOS) + 1;
		wsprintf(retbuf, "%d", nm); 
	} else if (!lstrcmpi(tag, "ratingstar")) {
		if (curitem->m_rating < 0)
			LoadMeta(curitem);

		lstrcpyn(retbuf, rateStars, curitem->m_rating+1);
	} else if (!lstrcmpi(tag, "666")) {
		wsprintf(retbuf, "Toasty!");
	} else if (!lstrcmpi(tag, "parentdir")) {
		char fn[MAX_PATH];
		lstrcpyn(fn, curitem->m_filename, MAX_PATH);
		for(int prf = lstrlen(fn); prf > 0; prf--)
		{
			if (fn[prf] == '\\')
			{
				fn[prf] = '\0';
				prf=0;
			}
		}
		for(prf = lstrlen(fn); prf > 0; prf--)
		{
			if (fn[prf] == '\\')
			{
				wsprintf(retbuf, "%s", &fn[prf+1]);
				prf=0;
			}
		}
	} else if (!lstrcmpi(tag, "dir")) {
		lstrcpyn(retbuf, curitem->m_filename, MAX_PATH);
		for(int prf = lstrlen(retbuf); prf > 0; prf--)
		{
			if (retbuf[prf] == '\\')
			{
				retbuf[prf] = '\0';
				prf=0;
			}
		}
	} else if (!lstrcmpi(tag, "ishttp")) {
		if (!strnicmp("http", curitem->m_filename, 4))
			wsprintf(retbuf, "Yes");
	} else if (!lstrcmpi(tag, "LENGTHF")) {
		//Handled by ML, but do it prettier	
		int s, m, h, length;
		if (curitem->m_length >= 0)
			length = curitem->m_length;
		else
		{
			if (LoadMeta(curitem))
			{
				efi.metadata = "LENGTH";
				SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
				curitem->m_length = atoi(retbuf);
				curitem->m_length = curitem->m_length / 1000;
			}
			length = curitem->m_length;
		}
		s = length % 60;
		m = (length / 60) % 60;
		h = (length / 60) / 60;
		if (h > 0) wsprintf(retbuf, "%.2d:%.2d:%.2d", h, m, s);
		else wsprintf(retbuf, "%.2d:%.2d", m, s);

	} else if (!lstrcmpi(tag, "BITRATE")) {
		if (!curitem->m_bmetaloaded && curitem->m_bitrate < 0)
		{
			if (LoadMeta(curitem))
			{
				efi.metadata = "BITRATE";
				SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
				curitem->m_bitrate = atoi(retbuf);
			}
		}
		wsprintf(retbuf, "%d", curitem->m_bitrate);
	} else if (!lstrcmpi(tag, "TITLE")) {
		if (!curitem->m_bmetaloaded && curitem->m_title[0] == '\0')
		{
			if (LoadMeta(curitem))
			{
				efi.metadata = "TITLE";
				efi.ret = curitem->m_title;
				SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
			}
		}
		strncpy(retbuf, curitem->m_title, MAX_PATH);
	} else if (!lstrcmpi(tag, "ARTIST")) {
		if (!curitem->m_bmetaloaded && curitem->m_artist[0] == '\0')
		{
			if (LoadMeta(curitem))
			{
				efi.metadata = "ARTIST";
				efi.ret = curitem->m_artist;
				SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
			}
		}
		strncpy(retbuf, curitem->m_artist, MAX_PATH);	}
	else if (!lstrcmpi(tag, "GENRE")) {
		if (!curitem->m_bmetaloaded && curitem->m_genre[0] == '\0')
		{
			if (LoadMeta(curitem))
			{
				efi.metadata = "GENRE";
				efi.ret = curitem->m_genre;
				SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
			}
		}
		strncpy(retbuf, curitem->m_genre, MAX_PATH);
	} else if (!lstrcmpi(tag, "ALBUM")) {
		if (!curitem->m_bmetaloaded && curitem->m_album[0] == '\0')
		{
			if (LoadMeta(curitem))
			{
				efi.metadata = "ALBUM";
				efi.ret = curitem->m_album;
				SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
			}
		}
		strncpy(retbuf, curitem->m_album, MAX_PATH);	}
	else if (!lstrcmpi(tag, "COMMENT")) {
			if (!curitem->m_bmetaloaded && curitem->m_comment[0] == '\0')
			{
				if (LoadMeta(curitem))
				{
					efi.metadata = "COMMENT";
					efi.ret = curitem->m_comment;
					SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
				}
			}
			strncpy(retbuf, curitem->m_comment, MAX_PATH);
	} else if (!lstrcmpi(tag, "YEAR")) {
		if (!curitem->m_bmetaloaded && curitem->m_year < 0)
		{
			if (LoadMeta(curitem))
			{
				efi.metadata = "YEAR";
				SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
				curitem->m_year = atoi(retbuf);
			}
		}
		wsprintf(retbuf, "%d", curitem->m_year);	
	} else if (!lstrcmpi(tag, "PLAYCOUNT")) {
		if (!curitem->m_bmetaloaded && curitem->m_playcount < 0)
			LoadMeta(curitem);

		wsprintf(retbuf, "%d", curitem->m_playcount);	
	} else if (!lstrcmpi(tag, "RATING")) {
		if (!curitem->m_bmetaloaded && curitem->m_rating < 0)
			LoadMeta(curitem);

		wsprintf(retbuf, "%d", curitem->m_rating);
	} else if (!lstrcmpi(tag, "LASTUPD")) {
		if (!curitem->m_bmetaloaded && curitem->m_lastupd < 0)
			LoadMeta(curitem);

		wsprintf(retbuf, "%d", curitem->m_lastupd);
	} else if (!lstrcmpi(tag, "LASTPLAY")) {
		if (!curitem->m_bmetaloaded && curitem->m_lastplay < 0)
			LoadMeta(curitem);

		wsprintf(retbuf, "%d", curitem->m_lastplay);
	} else if (!lstrcmpi(tag, "FILETIME")) {
		if (!curitem->m_bmetaloaded && curitem->m_filetime < 0)
			LoadMeta(curitem);

		wsprintf(retbuf, "%d", curitem->m_filetime);
	} else if (!lstrcmpi(tag, "FILESIZE")) {
		if (!curitem->m_bmetaloaded && curitem->m_filesize < 0)
			LoadMeta(curitem);

		wsprintf(retbuf, "%d", curitem->m_filesize);
	} else if (!lstrcmpi(tag, "LENGTH")) {
		if (!curitem->m_bmetaloaded && curitem->m_length < 0)
		{
			if (LoadMeta(curitem))
			{
				efi.metadata = "LENGTH";
				SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
				curitem->m_length = atoi(retbuf);
				curitem->m_length = curitem->m_length / 1000;
			}
		}
		wsprintf(retbuf, "%d", curitem->m_length);
	} else if (!lstrcmpi(tag, "shuffle")) {
		wsprintf(retbuf, "%d", SendMessage((HWND)p, WM_WA_IPC, 0,IPC_GET_SHUFFLE));
	} else if (!lstrcmpi(tag, "channelnum")) {
		wsprintf(retbuf, "%d", SendMessage((HWND)p, WM_WA_IPC, 2,IPC_GETINFO));
	} else if (!lstrcmpi(tag, "channels")) {
		lstrcpyn(retbuf, SendMessage((HWND)p, WM_WA_IPC, 2,IPC_GETINFO)-1?"stereo":"mono", MAX_PATH);
	} else if (!lstrcmpi(tag, "srate")) {
		wsprintf(retbuf, "%d", SendMessage((HWND)p, WM_WA_IPC, 0,IPC_GETINFO));
	} else if (!lstrcmpi(tag, "tracknumber")) {
		if (curitem->m_track < 0)
		{
			if (LoadMeta(curitem))
			{
				efi.metadata = "TRACK";
				SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
				curitem->m_track = atoi(retbuf);
				if (curitem->m_track == -1)
					curitem->m_track = 0;
			}
		}
		if (curitem->m_track >= 0)
			wsprintf(retbuf, "%.2d", curitem->m_track);
	} else if (!lstrcmpi(tag, "TRACK")) {
		if (curitem->m_track < 0)
		{
			if (LoadMeta(curitem))
			{
				efi.metadata = "TRACK";
				SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
				curitem->m_track = atoi(retbuf);
				if (curitem->m_track == -1)
					curitem->m_track = 0;
			}
		}
		if (curitem->m_track >= 0)
			wsprintf(retbuf, "%d", curitem->m_track);
	} else if (!strnicmp("cmp(", tag, 4))
	{
		nm=1;
		for (tp = 4; tp < lstrlen(tag) && nm > 0; tp++)
		{
			if (tag[tp] == '(') nm++;
			if (tag[tp] == ')') nm--;
		}
		tag[tp-1] = '\0';

		for (; tp > 4 && tag[tp] != '<' && tag[tp] != '>' && tag[tp] != '='; tp--);

		op[0] = tag[tp];
		op[1] = '\0';
		tag[tp] = '\0';

		tmpalloc = NULL;
		tmpalloc2 = NULL;
		ls = &tag[4];
		rs = &tag[tp+1];

		if (ls[0] == ':')
		{
			tmpalloc = myTagFunc(&ls[1], p);
			ls = tmpalloc;
		}
		if (rs[0] == ':')
		{
			tmpalloc2 = myTagFunc(&rs[1], p);
			rs = tmpalloc2;
		}
		if (ls != NULL && rs != NULL)
		{

			if (op[0] == '>')
			{
				nm = atoi(ls);
				nm2 = atoi(rs);

				if (nm > nm2)
					wsprintf(retbuf, ls);

			} else if (op[0] == '<')
			{
				nm = atoi(ls);
				nm2 = atoi(rs);
				if (nm < nm2)
					wsprintf(retbuf, ls);

			} else if (op[0] == '=')
			{
				if (!lstrcmpi(ls, rs))
					wsprintf(retbuf, ls);
			}
		}

		if (tmpalloc)
			myTagFreeFunc(tmpalloc, p);
		if (tmpalloc2)
			myTagFreeFunc(tmpalloc2, p);

	} else if (!strnicmp("tagln(", tag, 6))
	{
		nm=1;
		for (tp = 4; tp < lstrlen(tag) && nm > 0; tp++)
		{
			if (tag[tp] == '(') nm++;
			if (tag[tp] == ')') nm--;
		}
		tag[tp-1] = '\0';
		tmpalloc = NULL;
		ls = &tag[6];
		if (ls[0] == ':')
		{
			tmpalloc = myTagFunc(&ls[1], p);
			ls = tmpalloc;
		}

		if (ls != NULL)
			wsprintf(retbuf, "%d", lstrlen(ls));

		if (tmpalloc)
			myTagFreeFunc(tmpalloc, p);
	} else {
		//Try the media library
		char *tmp = NULL;
		mlQ.query=curitem->m_filename;
		mlQ.max_results=1;
		mlQ.results.Alloc=0;
		mlQ.results.Items=NULL;
		mlQ.results.Size=0;
		if (HwndMl && SendMessage(HwndMl,WM_ML_IPC,(WPARAM)&mlQ,ML_IPC_DB_RUNQUERY_FILENAME) == 1)
		{
			tmp = getRecordExtendedItem(&mlQ.results.Items[0], strupr(tag));
			if (tmp != NULL && lstrlen(tmp) != 0)
				lstrcpyn(retbuf, tmp, MAX_PATH);
		}

		if (HwndMl)
			SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlQ, ML_IPC_DB_FREEQUERYRESULTS);
		else
			HwndMl = (HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,0,(LPARAM)libmlipc);

		if (tmp == NULL)
		{
			//Else try winamp for the tag
			efi.metadata = tag;
			SendMessage((HWND)p, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
		}
	}

	if (lstrlen(retbuf))
	{
		s = (char*)GlobalAlloc(GPTR, 256);
		lstrcpyn(s, retbuf, 256);
		return s;
	}

	return 0;
}

STDMETHODIMP CMediaItem::get_Name(BSTR* pVal)
{
	USES_CONVERSION;
	*pVal = SysAllocString(T2OLE(m_filename));
	return S_OK;
}

int CMediaItem::CreateFromPlaylist(long index)
{
	char* fpt = (char *)SendMessage(plugin.hwndParent,WM_WA_IPC,index,IPC_GETPLAYLISTFILE);

	if (fpt != NULL)
	{
		strncpy(m_filename, fpt, sizeof(m_filename));
		m_plposition = index;
		return 0;
	} else
		return -1;
}

STDMETHODIMP CMediaItem::get_Filename(BSTR* pVal)
{
	// TODO: Add your implementation code here
	USES_CONVERSION;
	*pVal = SysAllocString(T2OLE(m_filename));
	return S_OK;
}

STDMETHODIMP CMediaItem::get_Position(LONG* pVal)
{
	*pVal = m_plposition + 1;
	return S_OK;
}

STDMETHODIMP CMediaItem::put_Position(LONG newVal)
{
	m_plposition = newVal - 1;
	return S_OK;
}

STDMETHODIMP CMediaItem::get_Title(BSTR* pVal)
{
	// TODO: Add your implementation code here
	USES_CONVERSION;
	extendedFileInfoStruct efi;

	if (m_title[0] != '\0')
		*pVal = SysAllocString(A2OLE(m_title));
	else {

		if (LoadMeta(this))
		{
			efi.filename = m_filename;
			efi.metadata = "title";
			efi.ret = m_title;
			efi.retlen = sizeof(m_title)-1;
			SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
		}
		*pVal = SysAllocString(A2OLE(m_title));
	}

	return S_OK;
}

STDMETHODIMP CMediaItem::put_Title(BSTR newVal)
{
	USES_CONVERSION;
	strncpy(m_title, OLE2T(newVal), sizeof(m_title));
	itemRecord mlR;
	memset(&mlR, 0, sizeof(itemRecord));
	mlR.filename=m_filename;
	mlR.title=m_title;
	mlR.year = -1;
	mlR.track = -1;
	mlR.length = -1;
	SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlR, ML_IPC_DB_UPDATEITEM);
	return S_OK;
}

STDMETHODIMP CMediaItem::ATFString(BSTR ATFSpecification, BSTR* FormattedString)
{
	USES_CONVERSION;
	char *pszString = OLE2A(ATFSpecification);
	char *tmp = GetFormattedTitleFromWinamp(pszString, this);
	*FormattedString = SysAllocString(A2OLE(tmp));
	GlobalFree(tmp);
	return S_OK;
}

int CMediaItem::CreateFromMLItem(itemRecord* iteminfo)
{
	if (iteminfo->filename != NULL)
	{
		strncpy(m_filename, iteminfo->filename, sizeof(m_filename));

		if (iteminfo->artist)
			strncpy(m_artist, iteminfo->artist, sizeof(m_artist));

		if (iteminfo->album)
			strncpy(m_album, iteminfo->album, sizeof(m_album));

		if (iteminfo->title)
			strncpy(m_title, iteminfo->title, sizeof(m_title));

		if (iteminfo->genre)
			strncpy(m_genre, iteminfo->genre, sizeof(m_genre));

		if (iteminfo->comment)
			strncpy(m_comment, iteminfo->comment, sizeof(m_comment));

		m_length = iteminfo->length;
		m_track = iteminfo->track;
		m_year = iteminfo->year;
		
		char *tmpc= getRecordExtendedItem(iteminfo, "RATING");
		if (tmpc != NULL)
			m_rating = tmpc[0] - '0';
		else
			m_rating = 0;

		tmpc = getRecordExtendedItem(iteminfo, "LASTPLAY");
		if (tmpc != NULL)
			m_lastplay = atol(tmpc);
		else
			m_lastplay = 0;

		tmpc = getRecordExtendedItem(iteminfo, "PLAYCOUNT");
		if (tmpc != NULL)
			m_playcount = atoi(tmpc);
		else
			m_playcount = 0;

		tmpc = getRecordExtendedItem(iteminfo, "BITRATE");
		if (tmpc != NULL)
			m_bitrate = atoi(tmpc);
		else
			m_bitrate = 0;

		tmpc = getRecordExtendedItem(iteminfo, "FILESIZE");
		if (tmpc != NULL)
			m_filesize = atoi(tmpc);
		else
			m_filesize = 0;

		tmpc = getRecordExtendedItem(iteminfo, "DBIDX");
		if (tmpc != NULL)
			m_dbidx = atoi(tmpc);
		else
			m_dbidx = 0;
		
		tmpc = getRecordExtendedItem(iteminfo, "LASTUPD");
		if (tmpc != NULL)
			m_lastupd = atol(tmpc);
		else
			m_lastupd = 0;

		tmpc = getRecordExtendedItem(iteminfo, "TYPE");
		if (tmpc != NULL)
			m_type = atoi(tmpc);
		else
			m_type = 0;

		tmpc = getRecordExtendedItem(iteminfo, "FILETIME");
		if (tmpc != NULL)
			m_filetime = atoi(tmpc);
		else
			m_filetime = 0;

		m_bmetaloaded = true;
		return 0;
	} else
		return -1;
}

STDMETHODIMP CMediaItem::Enqueue(void)
{
	COPYDATASTRUCT cds;

	//initialise the COPYDATASTRUCT
	cds.dwData = IPC_PLAYFILE;
	cds.lpData = (void*)m_filename;
	cds.cbData = (unsigned int)lstrlen((char*)cds.lpData) + 1; // Include space for null char at the end

	//Send the message to Winamp
	SendMessage(plugin.hwndParent,WM_COPYDATA, WPARAM(NULL),(LPARAM)&cds);
	return S_OK;
}

int CMediaItem::CreateFromFile(char* filename)
{
	if (filename != NULL)
	{
		strncpy(m_filename, filename, sizeof(m_filename));
		m_plposition = -1;
		return 0;
	} else
		return -1;
}

int CMediaItem::CreateFromFilePlaylist(char* filename, long index)
{
	if (filename != NULL)
	{
		strncpy(m_filename, filename, sizeof(m_filename));
		m_plposition = index;
		return 0;
	} else
		return -1;
}


STDMETHODIMP CMediaItem::get_Artist(BSTR* pVal)
{
	USES_CONVERSION;
	extendedFileInfoStruct efi;

	if (m_artist[0] != '\0')
		*pVal = SysAllocString(A2OLE(m_artist));
	else {
		if (LoadMeta(this))
		{
			efi.filename = m_filename;
			efi.metadata = "ARTIST";
			efi.ret = m_artist;
			efi.retlen = sizeof(m_artist)-1;
			SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
		}
		*pVal = SysAllocString(A2OLE(m_artist));
	}
	return S_OK;
}

STDMETHODIMP CMediaItem::put_Artist(BSTR newVal)
{
	// TODO: Add your implementation code here
	USES_CONVERSION;
	strncpy(m_artist, OLE2T(newVal), sizeof(m_artist));
	itemRecord mlR;
	memset(&mlR, 0, sizeof(itemRecord));
	mlR.filename=m_filename;
	mlR.artist=m_artist;
	mlR.year = -1;
	mlR.track = -1;
	mlR.length = -1;
	SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlR, ML_IPC_DB_UPDATEITEM);
	return S_OK;
}

STDMETHODIMP CMediaItem::get_Album(BSTR* pVal)
{
	USES_CONVERSION;
	extendedFileInfoStruct efi;

	if (m_album[0] != '\0')
		*pVal = SysAllocString(A2OLE(m_album));
	else {
		if (LoadMeta(this))
		{
			efi.filename = m_filename;
			efi.metadata = "ALBUM";
			efi.ret = m_album;
			efi.retlen = sizeof(m_album)-1;
			SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
		}
		*pVal = SysAllocString(A2OLE(m_album));
	}

	return S_OK;
}

STDMETHODIMP CMediaItem::put_Album(BSTR newVal)
{
	// TODO: Add your implementation code here
	USES_CONVERSION;
	strncpy(m_album, OLE2T(newVal), sizeof(m_album));
	itemRecord mlR;
	memset(&mlR, 0, sizeof(itemRecord));
	mlR.filename=m_filename;
	mlR.album=m_album;
	mlR.year = -1;
	mlR.track = -1;
	mlR.length = -1;
	SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlR, ML_IPC_DB_UPDATEITEM);
	return S_OK;
}

STDMETHODIMP CMediaItem::get_Rating(BYTE* pVal)
{
	if (m_rating >= 0)
		*pVal = (BYTE)m_rating;
	else
	{
		LoadMeta(this);
		*pVal = (BYTE)m_rating;
	}
	
	return S_OK;
}

STDMETHODIMP CMediaItem::put_Rating(BYTE newVal)
{
	setItemVal(m_filename, "RATING", newVal);
	m_rating = newVal;
	return S_OK;
}

STDMETHODIMP CMediaItem::get_Playcount(LONG* pVal)
{
	if (m_playcount > 0)
		*pVal = m_playcount;
	else
	{
		LoadMeta(this);
		*pVal = m_playcount;
	}
	return S_OK;
}

STDMETHODIMP CMediaItem::put_Playcount(LONG newVal)
{
	setItemVal(m_filename, "PLAYCOUNT", newVal);
	m_playcount = newVal;
	return S_OK;
}

STDMETHODIMP CMediaItem::Insert(LONG index)
{
	COPYDATASTRUCT cds;
	
	fileinfo finf;
	strncpy(finf.file, m_filename, MAX_PATH);
	finf.index = index;

	//initialise the COPYDATASTRUCT
	cds.dwData = IPC_PE_INSERTFILENAME;
	cds.lpData = (void*)&finf;
	cds.cbData = sizeof(finf); // Include space for null char at the end

	//Send the message to Winamp
	SendMessage(HwndPl,WM_COPYDATA, WPARAM(NULL),(LPARAM)&cds);

	return S_OK;
}

STDMETHODIMP CMediaItem::get_LastPlay(LONG* pVal)
{
	if (m_lastplay > 0)
		*pVal = m_lastplay;
	else
	{
		LoadMeta(this);
		*pVal = m_lastplay;
	}
	return S_OK;
}

STDMETHODIMP CMediaItem::put_LastPlay(LONG newVal)
{
	setItemVal(m_filename, "LASTPLAY", newVal);
	m_lastplay = newVal;
	return S_OK;
}


int LoadMeta(CMediaItem* mediai)
{
	mlQueryStruct mlQ;
	mlQ.query=mediai->m_filename;
	mlQ.max_results=1;
	mlQ.results.Alloc=0;
	mlQ.results.Items=NULL;
	mlQ.results.Size=0;

	if (SendMessage(HwndMl,WM_ML_IPC,(WPARAM)&mlQ,ML_IPC_DB_RUNQUERY_FILENAME) == 1)
	{
		mediai->CreateFromMLItem(&mlQ.results.Items[0]);
		SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlQ, ML_IPC_DB_FREEQUERYRESULTS);
		return 0;
	}
	return -1;
}

STDMETHODIMP CMediaItem::RefreshMeta(void)
{
	// TODO: Add your implementation code here
	char retbuf[MAX_PATH];
	mlQueryStruct mlQ;
	extendedFileInfoStruct efi;

	mlQ.query=m_filename;
	mlQ.max_results=1;
	mlQ.results.Alloc=0;
	mlQ.results.Items=NULL;
	mlQ.results.Size=0;

	if (SendMessage(HwndMl,WM_ML_IPC,(WPARAM)&mlQ,ML_IPC_DB_RUNQUERY_FILENAME) == 1)
	{
		CreateFromMLItem(&mlQ.results.Items[0]);
		SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlQ, ML_IPC_DB_FREEQUERYRESULTS);
	} else
	{
		ZeroMemory(retbuf, MAX_PATH);
		efi.filename = m_filename;
		efi.retlen = sizeof(retbuf)-1;

		efi.metadata = "ARTIST";
		efi.ret = m_artist;
		SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
				
		efi.metadata = "TITLE";
		efi.ret = m_title;
		SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);

		efi.metadata = "GENRE";
		efi.ret = m_genre;
		SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
	
		efi.metadata = "COMMENT";
		efi.ret = m_comment;
		SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);

		efi.metadata = "ALBUM";
		efi.ret = m_album;
		SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
		
		efi.metadata = "LENGTH";
		efi.ret = retbuf;
		SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
		m_length = atoi(retbuf);
		m_length = m_length / 1000;

		efi.metadata = "BITRATE";
		SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
		m_bitrate = atoi(retbuf);

		efi.metadata = "TRACK";
		SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
		m_track = atoi(retbuf);
        if (m_track == -1)
			m_track = 0;

		efi.metadata = "YEAR";
		SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
		m_track = atoi(retbuf);

	}
	m_bmetaloaded = true;
	return S_OK;
}

STDMETHODIMP CMediaItem::get_DbIndex(LONG* pVal)
{
	if (m_dbidx >= 0)
		*pVal = m_dbidx;
	else
	{
		LoadMeta(this);
		*pVal = m_dbidx;
	}
	return S_OK;
}

STDMETHODIMP CMediaItem::get_Length(LONG *pVal)
{
	extendedFileInfoStruct efi;
	char retbuf[10];

	if (m_length >= 0)
		*pVal = m_length;
	else
	{
		if (LoadMeta(this))
		{
			efi.filename = m_filename;
			efi.retlen = 10;
			efi.metadata = "LENGTH";
			efi.ret = retbuf;
			SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
			m_length = atoi(retbuf);
			m_length = m_length / 1000;
		}
		*pVal = m_length;
	}
	return S_OK;
}

STDMETHODIMP CMediaItem::get_Track(LONG* pVal)
{
	extendedFileInfoStruct efi;
	char retbuf[10];

	if (m_track >= 0)
		*pVal = m_track;
	else
	{
		if (LoadMeta(this))
		{
			efi.filename = m_filename;
			efi.retlen = 10;
			efi.metadata = "TRACK";
			efi.ret = retbuf;
			SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
			m_track = atoi(retbuf);
		}
		*pVal = m_track;
	}
	return S_OK;
}

STDMETHODIMP CMediaItem::get_Genre(BSTR* pVal)
{
	USES_CONVERSION;
	extendedFileInfoStruct efi;

	if (m_genre[0] != '\0')
		*pVal = SysAllocString(A2OLE(m_genre));
	else {
		if (LoadMeta(this))
		{
			efi.filename = m_filename;
			efi.metadata = "GENRE";
			efi.ret = m_genre;
			efi.retlen = sizeof(m_genre)-1;
			SendMessage((HWND)plugin.hwndParent, WM_WA_IPC, (WPARAM)&efi, IPC_GET_EXTENDED_FILE_INFO_HOOKABLE);
		}
		*pVal = SysAllocString(A2OLE(m_genre));
	}
	return S_OK;
}

STDMETHODIMP CMediaItem::put_Genre(BSTR newVal)
{
	USES_CONVERSION;
	strncpy(m_genre, OLE2T(newVal), sizeof(m_genre));
	itemRecord mlR;
	memset(&mlR, 0, sizeof(itemRecord));
	mlR.filename=m_filename;
	mlR.genre=m_genre;
	mlR.year = -1;
	mlR.track = -1;
	mlR.length = -1;
	SendMessage(HwndMl, WM_ML_IPC, (WPARAM) &mlR, ML_IPC_DB_UPDATEITEM);
	return S_OK;
}