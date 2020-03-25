/************************************************************************

ScriptManager.cpp

Used mainly for the management of internal script hosting. Also contains
the hooks for the main window. playlist and media library.

TODO: 
Fix thread script thread pooling and 'quitting' of scripts
Use of "Clone" to duplicate the script host instead of re-creating it
Better 'running scripts' GUI support

************************************************************************/

#include "StdAfx.h"
#include <map>
#include <string>
#include <list>

#include <initguid.h>
#include ".\scriptmanager.h"
DEFINE_GUID(CLSID_VBScript, 0xb54f3741, 0x5b07, 0x11cf, 0xa4, 0xb0, 0x0,0xaa, 0x0, 0x4a, 0x55, 0xe8);

extern CApplication* pApplication;

HMENU actionmenu;
long spareMenuId=61666;
bool bMenuAdded;
long lastmenuid=0;
long lastsendid=0;
typedef std::pair <long, std::string> Pls_Pair;
typedef std::map <long, std::string> SCRIPTIDS;
SCRIPTIDS plscommandmap;
SCRIPTIDS sendtocommandmap;
typedef std::list <CMyScriptSite*> LISTSCRIPT;
LISTSCRIPT runningScripts;
ITypeLib *ptLib = 0;
bool bcleared = false;
extern CScriptManager glbScriptManager;
extern DWORD pappCookie;

void HRVERIFY(HRESULT hr, char * msg)
{
	if(FAILED(hr)) {
		static char buf[1024];
		wsprintf(buf, "Error: %08lx, %s", hr, msg);
		::MessageBox(NULL, buf, "", MB_SETFOREGROUND);
		//_exit(0);
	}

}

LRESULT CALLBACK MWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	//SiteManager.quit effectively sends this message.
	if (message == quit_ipc)
	{
		LISTSCRIPT::iterator si;
		for (si = runningScripts.begin(); si != runningScripts.end(); ++si)
		{
			if (*si == (CMyScriptSite*)wParam)
			{
				//Interrupt the script and cause it to quit
				HRESULT	hr = ((CMyScriptSite*)wParam)->m_scriptp->InterruptScriptThread(SCRIPTTHREADID_ALL, NULL, SCRIPTINTERRUPT_RAISEEXCEPTION);
				//HANDLE threadhandle = OpenThread(SYNCHRONIZE, false, ((CMyScriptSite*)wParam)->threadid);
				
				//Tell the workerthread of the desired script to quit
				PostThreadMessage(((CMyScriptSite*)wParam)->threadid, WM_QUIT, 0, 0);
				//CloseHandle(threadhandle);
				//runningScripts.remove(*si);

				//There is something with this quitting process which is causing threads to not be
				//cleaned up properly when running many scripts fast in succession
				break;
			}
		}
	}
	else if (message == hotkey_ipc)
	{
		std::map <long, std::string> :: iterator hotIter;
		hotIter = plscommandmap.find(lParam);
		if ( hotIter != plscommandmap.end() )
			glbScriptManager.runfile(hotIter->second.c_str(), true, "");
	}
	else if (message == init_ipc)
	{
		glbScriptManager.init();
		glbScriptManager.run();
	} else if (message == WM_WA_IPC)
	{
		if (lParam == IPC_HOOK_OKTOQUIT)
		{
			int retval = CallWindowProc(lpWndProcOld,hwnd,message,wParam,lParam);
			if (retval != 0)
			{
				mlplugin.hDllInstance = NULL;
				SendMessage(HwndMl, WM_ML_IPC, (WPARAM)&mlplugin, ML_IPC_REMOVE_PLUGIN);
			} 
			return retval;
		}
		
		if (lParam==IPC_CB_MISC)
		{
			switch(wParam)
			{
			case IPC_CB_MISC_VOLUME:
				if (pApplication)
					pApplication->Fire_ChangedVolume();
				break;
			case IPC_CB_MISC_STATUS:
				if (pApplication)
					pApplication->Fire_ChangedStatus();
				break;
			case IPC_CB_MISC_INFO:
			case IPC_CB_MISC_VIDEOINFO:
			case IPC_CB_MISC_TITLE:
				break;
			}
		}
	} else if (message == WM_WA_MPEG_EOF)
	{
		if (pApplication)
			pApplication->Fire_PlaybackEOF();
	}
	return CallWindowProc(lpWndProcOld,hwnd,message,wParam,lParam);
}

int MLWndProc(int message, int param1, int param2, int param3)
{
	mlAddToSendToStruct mlsendadd;
	std::map <long, std::string> :: iterator sendIter;
	long val;
	char desc[MAX_PATH];

	switch (message)
	{
	case ML_MSG_ONSENDTOBUILD:
		mlsendadd.context = param2;
		for ( sendIter = sendtocommandmap.begin( ) ; sendIter != sendtocommandmap.end( ); sendIter++)
		{
			wsprintf(desc, "Script: %s", (char*)sendIter->second.c_str()+7);
			mlsendadd.desc = desc;
			strtok(mlsendadd.desc, ".");
			//mlsendadd.desc = (char*)sendIter->second.c_str();
			mlsendadd.user32 = sendIter->first + (long)MLWndProc;
			SendMessage(HwndMl, WM_ML_IPC, (WPARAM)&mlsendadd, ML_IPC_ADDTOSENDTO);
		}
		break;
	case ML_MSG_ONSENDTOSELECT:
		if (param3 >= (long)MLWndProc && param3 <= (long)MLWndProc + lastsendid)
		{
			pApplication->CreateSendToItems(param1, (void*)param2);
			sendIter = sendtocommandmap.find(param3-(long)MLWndProc);
			if ( sendIter != sendtocommandmap.end() )
				glbScriptManager.runfile(sendIter->second.c_str(), true, "");
		}
		break;

	}
	return 0;
}

LRESULT CALLBACK PlWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_USER:
		if (wParam == 666)
		{
			if ((lParam & 0x40000000) == 0x40000000)
			{
				if (bcleared)
				{
					if (pApplication)
						pApplication->Fire_ChangedTrack();

					char *scp = (strstr((char*)SendMessage(plugin.hwndParent, WM_WA_IPC, (lParam & 0x0FFFFFFF), IPC_GETPLAYLISTFILE), "script:\\"));
					if (scp)
					{
						char argbuf[MAX_PATH];
						char *sep = strtok(scp+8, " ");
						sep = strtok(NULL, " ");
						wsprintf(argbuf, "-pos=%d %s", (lParam & 0x0FFFFFFF)+1, sep);
						glbScriptManager.runfile(scp+8, true, argbuf);
					}
				}
				bcleared = false;
			} else
				bcleared = true;
		}
		break;

	case WM_INITMENUPOPUP:
		MENUITEMINFO mii;
		char buf[64];

		mii.cbSize = sizeof(MENUITEMINFO);
		mii.fMask = MIIM_ID|MIIM_TYPE|MIIM_DATA|MIIM_SUBMENU;
		mii.dwTypeData = buf;
		mii.cch = sizeof(buf);
		GetMenuItemInfo((HMENU)wParam,0,true,&mii);
		//int val = GetMenuItemID((HMENU)wParam,0);

		//"Play Item/tEnter"
		if (mii.wID==0x9cf8)
		{
			mii.wID = spareMenuId;
			if (bMenuAdded == false)
			{
				mii.fType = MFT_STRING;
				lstrcpyn(buf,"Scripts",sizeof(buf));
				mii.hSubMenu=actionmenu;
				InsertMenuItem((HMENU)wParam,9,1,&mii);
				bMenuAdded = true;
			}
		}
		break;
	case WM_COMMAND:
		if ((LOWORD(wParam) > spareMenuId) && (LOWORD(wParam) <= lastmenuid))
		{
			std::map <long, std::string> :: iterator plsIter;

			plsIter = plscommandmap.find((long)wParam);
			if ( plsIter != plscommandmap.end() )
				glbScriptManager.runfile(plsIter->second.c_str(), true, "");
		}
		break;
	}
	return CallWindowProc(lpPlWndProcOld,hwnd,message,wParam,lParam);
}

//Called at init time, other plugins have not necessarily initted at this point.
//The main COM object is also not registered in the active table.
int CScriptManager::init1(void)
{
	lstrcpyn(iniDir, (LPCSTR)SendMessage(plugin.hwndParent,WM_WA_IPC,0,IPC_GETINIFILE), MAX_PATH);

	int fnl = lstrlen(iniDir);
	for(int prf = fnl; prf > 0; prf--)
	{
		if (iniDir[prf] == '\\')
		{
			iniDir[prf] = '\0';
			prf = 0;
		}
	}
	wsprintf(iniDir, "%s%s", iniDir, "\\Plugins\\Scripts\\");

	MENUITEMINFO mii2;
	//Setup menus
	actionmenu = CreateMenu();

	mii2.cbSize = sizeof(MENUITEMINFO);
	mii2.fMask = MIIM_ID|MIIM_TYPE|MIIM_DATA;
	mii2.cch=MAX_PATH;
	mii2.wID=spareMenuId+1;
	mii2.fType=MFT_STRING;

	WIN32_FIND_DATA FindFileData;
	HANDLE hFind = INVALID_HANDLE_VALUE;
	HANDLE hFile;
	DWORD bytesread;
	char fn[MAX_PATH];

	wsprintf(fn, "%s%s", iniDir, "playlist_*.vbs");
	bool finished = false;
	hFind = FindFirstFile(fn, &FindFileData);
	while (hFind != INVALID_HANDLE_VALUE && !finished)
	{
		std::string plscom (FindFileData.cFileName);
		plscommandmap.insert ( Pls_Pair ( mii2.wID, plscom ) );

		//Menu item
		mii2.dwTypeData = FindFileData.cFileName+9;
		strtok(mii2.dwTypeData, ".");
		InsertMenuItem(actionmenu, mii2.wID-spareMenuId, 1, &mii2);
		lastmenuid=mii2.wID;

		mii2.wID++;
		if (!FindNextFile(hFind, &FindFileData))
			finished = true;
	}

	wsprintf(fn, "%s%s", iniDir, "sendto_*.vbs");
	finished = false;
	hFind = FindFirstFile(fn, &FindFileData);
	int idx = 0;
	while (hFind != INVALID_HANDLE_VALUE && !finished)
	{
		std::string sendfn (FindFileData.cFileName);
		sendtocommandmap.insert ( Pls_Pair ( idx, sendfn ) );
		lastsendid = idx;	
		idx++;
		if (!FindNextFile(hFind, &FindFileData))
			finished = true;
	}

	return 0;
}


//Called after other plugins have initialised.
void CScriptManager::init()
{
	//Register the Automation object ready to use
	RegisterActiveObject(pApplication->GetUnknown(), CLSID_Application, ACTIVEOBJECT_STRONG, &dwr);
	//IRunnableObject *pRot;
	//pRot = GetRunningObjectTable(0, &pRot);
	//pRot->Register()
	CoLockObjectExternal(pApplication->GetUnknown(),true,true);

	//Get Media Library IPC and HWND
	libmlipc = (LONG)SendMessage(plugin.hwndParent, WM_WA_IPC, (WPARAM)"LibraryGetWnd", IPC_REGISTER_WINAMP_IPCMESSAGE);

	char fn[MAX_PATH];

	long genhotkeys_add_ipc=SendMessage(plugin.hwndParent,WM_WA_IPC,(WPARAM)"GenHotkeysAdd",IPC_REGISTER_WINAMP_IPCMESSAGE);	
	genHotkeysAddStruct ghks;
	SCRIPTIDS::iterator si;
	for (si = plscommandmap.begin(); si != plscommandmap.end(); si++)
	{
		//Hotkey item
		wsprintf(fn, "%s%s", "Script: ", std::string((*si).second).c_str());
		ghks.wnd=0;
		ZeroMemory(ghks.extended, sizeof(ghks.extended));
		ghks.flags=0;
		ghks.id=fn;
		ghks.uMsg=hotkey_ipc;
		ghks.lParam=(*si).first;
		ghks.wParam=0;
		ghks.name=fn;
		SendMessage(plugin.hwndParent,WM_WA_IPC,(WPARAM)&ghks,genhotkeys_add_ipc);
	}
	//Add Media Library stub program
	mlplugin.hDllInstance = NULL;
	if (!HwndMl)
		HwndMl = (HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,-1,(LPARAM)libmlipc);

	SendMessage(HwndMl, WM_ML_IPC, (WPARAM)&mlplugin, ML_IPC_ADD_PLUGIN);
}

struct threadparam {
	char *fn;
	bool fromdir;
	std::string args;
};

//Runs an internal script.
//Should be modified to re-use script host objects
//through .Clone etc. Also script 'watchdog'ing needs
//to be fixed.
DWORD WINAPI WorkerThreadFunc(LPVOID Param)
{
	USES_CONVERSION;

	CoInitialize(NULL);
	///CoInitializeEx(NULL, COINIT_MULTITHREADED );

	threadparam *tp = (threadparam*)Param;

	CMyScriptSite *pMySite;
	CSiteManager  *pSiteManager;
	IActiveScript *pAS;
	IActiveScriptParse *pASP;
	char fullfn[MAX_PATH];
	HANDLE hFile;
	DWORD bytesread;

	if (tp->fromdir)
		wsprintf(fullfn, "%s%s", glbScriptManager.iniDir, tp->fn);
	else
		wsprintf(fullfn, "%s", tp->fn);

	hFile = CreateFile(fullfn, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);

	if (hFile != INVALID_HANDLE_VALUE)
	{

	unsigned numChars = GetFileSize(hFile, 0);
	// create a file-mapping object for the script file
	HANDLE hMap = CreateFileMapping(hFile, 0, PAGE_READONLY, 0, 0, 0);
	char *pScriptAscii = (char *) MapViewOfFile(hMap, FILE_MAP_READ, 0, 0, 0);
	// create a BSTR to hold the translated text
	CComBSTR script(numChars);
	// translate ASCII to Unicode
	mbstowcs(script.m_str, pScriptAscii, numChars);

	UnmapViewOfFile(pScriptAscii);
	pScriptAscii = 0;
	CloseHandle(hMap);
	hMap = 0;
	CloseHandle(hFile);
	hFile = INVALID_HANDLE_VALUE;

	pMySite = new CMyScriptSite;
	pSiteManager = new CComObject<CSiteManager>;
	pSiteManager->m_hostSite = pMySite;
	pMySite->m_sitemanage = pSiteManager;

	pMySite->m_desc = tp->fn;
	pMySite->m_arguments = &tp->args;
	runningScripts.push_back(pMySite);

	// Register your type-library
	if (ptLib == NULL)
	{
		TCHAR szPath[MAX_PATH];
		GetModuleFileName(mhin, szPath, MAX_PATH);
		HRVERIFY(LoadTypeLib(T2COLE(szPath), &ptLib), "LoadTypeLib");
	}

	// Initialize your IActiveScriptSite implementation with your
	// object's ITypeInfo...
	ptLib->GetTypeInfoOfGuid(CLSID_Application, &pMySite->m_pTypeInfo);
	//ptLib->Release();

	// Initialize your IActiveScriptSite implementation with your
	// script object's IUnknown interface...
	
	IGlobalInterfaceTable *gpGIT;
	IApplication *pApp;
	CoCreateInstance(CLSID_StdGlobalInterfaceTable,NULL,CLSCTX_INPROC_SERVER,IID_IGlobalInterfaceTable,(void **)&gpGIT);
	gpGIT->GetInterfaceFromGlobal(pappCookie, IID_IApplication,(void**)&pApp);
	gpGIT->Release();
	
	HRVERIFY(pApp->QueryInterface(IID_IUnknown,
		(void **)&pMySite->m_pUnkScriptObject), "Script Object IUnknown initialization");

	HRVERIFY(pSiteManager->QueryInterface(IID_IUnknown,
		(void **)&pMySite->m_pUnkManagerObject), "Manager IUnknown initialization");

	// Start inproc script engine, VBSCRIPT.DLL

	HRVERIFY(CoCreateInstance(CLSID_VBScript, NULL, CLSCTX_INPROC_SERVER,
		IID_IActiveScript, (void **)&pAS), 
		"CoCreateInstance() for CLSID_VBScript");

	// Get engine's IActiveScriptParse interface.
	HRVERIFY(pAS->QueryInterface(IID_IActiveScriptParse, (void **)&pASP),
		"QueryInterface() for IID_IActiveScriptParse");

	pMySite->m_parse = pASP;
	pMySite->m_scriptp=pAS;

	// Give engine your IActiveScriptSite interface...
	HRVERIFY(pAS->SetScriptSite((IActiveScriptSite *)pMySite),
		"IActiveScript::SetScriptSite()");

	// Give the engine a chance to initialize itself...
	HRVERIFY(pASP->InitNew(), "IActiveScriptParse::InitNew()");

	// Add a root-level item to the engine's name space...
	HRVERIFY(pAS->AddNamedItem(L"Application", SCRIPTITEM_ISVISIBLE |
		SCRIPTITEM_ISSOURCE | SCRIPTITEM_GLOBALMEMBERS), "IActiveScript::AddNamedItem()");

	HRVERIFY(pAS->AddNamedItem(L"SiteManager", SCRIPTITEM_ISVISIBLE  | SCRIPTITEM_GLOBALMEMBERS)
	 , "IActiveScript::AddNamedItem()");

	// Parse the code scriptlet...
	EXCEPINFO ei;
	if (pASP->ParseScriptText(script.m_str, L"Application", NULL, NULL, 0, 0, 0L, NULL, &ei) == DISP_E_EXCEPTION)
	{
		SysFreeString(ei.bstrDescription);
		SysFreeString(ei.bstrHelpFile);
		SysFreeString(ei.bstrSource);
	} else
	{
		// Set the engine state. This line actually triggers the execution
		// of the script.
		HRVERIFY(pAS->SetScriptState(SCRIPTSTATE_CONNECTED), "SetScriptState");

		MSG msg;
		//Wait for events until script quits.
		//The myscriptsite sends a WM_QUIT to this thread when it has been terminated.
		while (GetMessage(&msg, 0, 0, 0))
		{
			DispatchMessage(&msg);
		}
	}
	
	pApp->Release();
	//Bad thread pooling implementation
	runningScripts.remove(pMySite);

	delete pMySite;

	}
	delete tp->fn;
	delete tp;
	
	CoUninitialize();
	return 0;
}

void CScriptManager::runfile(const char *fn, bool fromDir, const char *args)
{
	if (pApplication)
	{
		DWORD ThreadId;
		threadparam *tp = new threadparam;
		tp->fn = new char[MAX_PATH];
		strncpy(tp->fn, fn, MAX_PATH);
		tp->fromdir = fromDir;
		tp->args = args;
		CreateThread(NULL,0,WorkerThreadFunc,(LPVOID)tp,0, &ThreadId);
	}
}


void CScriptManager::run()
{
	WIN32_FIND_DATA FindFileData;
	HANDLE hFind = INVALID_HANDLE_VALUE;
	
	char fn[MAX_PATH];
	wsprintf(fn, "%s%s", iniDir, "startup_*.vbs");

	bool finished = false;
	hFind = FindFirstFile(fn, &FindFileData);

	while (hFind != INVALID_HANDLE_VALUE && !finished)
	{
		runfile(FindFileData.cFileName, true, "");

		if (!FindNextFile(hFind, &FindFileData))
			finished = true;
	}
	FindClose(hFind);
}

void CScriptManager::stopall(void)
{
	LISTSCRIPT::iterator si;
	si = runningScripts.begin();

	while (si != runningScripts.end()) {
		CMyScriptSite *tmp = *si;
		HANDLE threadhandle = OpenThread(SYNCHRONIZE, false, tmp->threadid);
		HRESULT	hr = tmp->m_scriptp->InterruptScriptThread(SCRIPTTHREADID_ALL, NULL, SCRIPTINTERRUPT_RAISEEXCEPTION);
		PostThreadMessage(tmp->threadid, WM_QUIT, 0, 0);

		//Wait for the thread to exit, processing window messages so that COM can
		//properly let the script quit.
		while (TRUE)
		{
			// block-local variable 
			DWORD result ; 
			MSG msg ; 
			// Read all of the messages in this next loop, 
			// removing each message as we read it.
			while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) 
			{ 
				// Otherwise, dispatch the message.
				DispatchMessage(&msg); 
			} // End of PeekMessage while loop.

			// Wait for any message sent or posted to this queue 
			// or for one of the passed handles be set to signaled.
			result = MsgWaitForMultipleObjects(1, &threadhandle, FALSE, INFINITE, QS_ALLINPUT);

			// The result tells us the type of event we have.
			if (result == (WAIT_OBJECT_0 + 1))
			{
				// New messages have arrived. 
				// Continue to the top of the always while loop to 
				// dispatch them and resume waiting.
				continue;
			} 
			else 
			{ 
				// One of the handles became signaled. 
				break;
			} // End of else clause.
		} // End of the always while loop. 
		CloseHandle(threadhandle);
		si = runningScripts.begin();
	};
}

CScriptManager::CScriptManager(void)
{
}

CScriptManager::~CScriptManager(void)
{
	if (ptLib)
		ptLib->Release();

	if (pApplication)
	{
		IGlobalInterfaceTable *gpGIT;
		CoCreateInstance(CLSID_StdGlobalInterfaceTable,	NULL,CLSCTX_INPROC_SERVER,IID_IGlobalInterfaceTable,(void **)&gpGIT);
		gpGIT->RevokeInterfaceFromGlobal(pappCookie);
		gpGIT->Release();
		AtlComModuleRevokeClassObjects(&ATL::_AtlComModule);
		RevokeActiveObject(dwr, NULL);
		CoLockObjectExternal(pApplication->GetUnknown(), false, true);
		CoDisconnectObject(pApplication->GetUnknown(), 0);
		pApplication->Release();
	}
}
