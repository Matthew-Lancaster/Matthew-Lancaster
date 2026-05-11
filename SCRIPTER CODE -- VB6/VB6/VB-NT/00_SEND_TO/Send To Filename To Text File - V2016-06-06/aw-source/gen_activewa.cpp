/*************************************************************************
*
* ActiveWinamp
*
* An automation plugin for Winamp. This exposes a lot of winamp functionality 
* through COM / ActiveX. You may use this to write VB or .NET or any other 
* applications which interact with Winamp.
* 
* You can automate Winamp internally or externally. Write scripts that can 
* be executed from the playlist, "Send To" menu or hot keys, or from external 
* scripts and programs. 
*
* Official Website: https://sourceforge.net/projects/activewinamp/
* Winamp Plugin Site: http://www.winamp.com/plugins/details.php?id=143299
*
*-------------------------------------------------------------------------
*
* Name    :	gen_activewa.cpp
*
* Purpose : Implementation of DLL Exports and object registration
*
* Author  : Shane Hird
*
* License : GNU Library or Lesser General Public License (LGPL)
*
************************************************************************/ 

#include "stdafx.h"
#include "resource.h"
#include "gen_activewa.h"
#include "globals.h"
#include "ipc_pe.h"
#include "RunningList.h"

#include "ScriptManager.h"
#include "Application.h"

CScriptManager glbScriptManager;
CApplication *pApplication;
CComObject<CApplication> *pApp;

class Cgen_activewaModule : public CAtlDllModuleT< Cgen_activewaModule >
{
public :
DECLARE_LIBID(LIBID_ActiveWinamp)
};

Cgen_activewaModule _AtlModule;

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	mhin = hInstance;
    //return _AtlModule.DllMain(dwReason, lpReserved); 
	return 1;
}

CRunningList rl;

void config()
{
	rl.Create(plugin.hwndParent);
	rl.ShowWindow(SW_SHOW);
}

void quit()
{
	glbScriptManager.stopall();
	SetWindowLong(HwndPl, GWL_WNDPROC, (LONG)lpPlWndProcOld);
	SetWindowLong(plugin.hwndParent, GWL_WNDPROC, (LONG)lpWndProcOld);
}

void quitML()
{
	mlplugin.hDllInstance = NULL;
}


//Assumes you are running regsvr32 from the plugins directory.
void RegisterAdditionalResource(UINT ResId)
{
	USES_CONVERSION;
	HRESULT hRes = S_OK;
	CComPtr<IRegistrar> p;

	hRes = CoCreateInstance(CLSID_Registrar, NULL,
		CLSCTX_INPROC_SERVER, IID_IRegistrar, (void**)&p);

	if (SUCCEEDED(hRes))
	{
		//HINSTANCE hWinamp = (HINSTANCE)GetWindowLong(plugin.hwndParent, GWL_HINSTANCE); 
		TCHAR szModule[_MAX_PATH];
	
		GetModuleFileName(mhin, szModule, _MAX_PATH);
		
		LPOLESTR pszModuleAW = T2OLE(szModule);
		int fnl = lstrlen(szModule);

		int pc = 2;
		for(int prf = fnl; prf > 0 && pc > 0; prf--)
		{
			if (szModule[prf] == '\\')
			{
				szModule[prf] = '\0';
				pc--;
			}
		}
		wsprintf(szModule, "%s%s", szModule, "\\winamp.exe");

		LPOLESTR pszModule = T2OLE(szModule);
		p->AddReplacement(OLESTR("Module"), pszModule);
		LPOLESTR szType = OLESTR("REGISTRY");
		hRes = p->ResourceRegister(pszModuleAW, ResId, szType);
	}
}


int init()
{
	_AtlModule.m_libid = LIBID_ActiveWinamp;

	HwndPl=(HWND)SendMessage(plugin.hwndParent,WM_WA_IPC,IPC_GETWND_PE,IPC_GETWND);
	lpPlWndProcOld = (WNDPROC)SetWindowLong(HwndPl, GWL_WNDPROC, (LONG)PlWndProc);

	lpWndProcOld = (WNDPROC)SetWindowLong(plugin.hwndParent, GWL_WNDPROC, (LONG)MWndProc);

	quit_ipc=SendMessage(plugin.hwndParent,WM_WA_IPC,(WPARAM)"ScriptQuitIPC",IPC_REGISTER_WINAMP_IPCMESSAGE);	
	hotkey_ipc=SendMessage(plugin.hwndParent,WM_WA_IPC,(WPARAM)"ScriptHotKey",IPC_REGISTER_WINAMP_IPCMESSAGE);	
	init_ipc=SendMessage(plugin.hwndParent,WM_WA_IPC,(WPARAM)"ScriptInit",IPC_REGISTER_WINAMP_IPCMESSAGE);	

	//pApp = new CComObject<CApplication>;
	//pApplication = pApp;
	//IApplication* ppv2;
	//CoCreateInstance(CLSID_Application,NULL,CLSCTX_INPROC_SERVER,IID_IApplication,(LPVOID*)&ppv2);
	//pApplication = (CApplication*)ppv2;
	AtlComModuleRegisterClassObjects(&ATL::_AtlComModule, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE);

	IApplication* ppv2;
	CoCreateInstance(CLSID_Application,NULL,CLSCTX_INPROC_SERVER,IID_IApplication,(LPVOID*)&ppv2);
	pApplication = (CApplication*)ppv2;
	
	IGlobalInterfaceTable *gpGIT;
	CoCreateInstance(CLSID_StdGlobalInterfaceTable,	NULL, CLSCTX_INPROC_SERVER,	IID_IGlobalInterfaceTable,(void **)&gpGIT);
	gpGIT->RegisterInterfaceInGlobal(pApplication->GetUnknown(), IID_IApplication, &pappCookie);
	gpGIT->Release();

	glbScriptManager.init1();
	//Post a message to init after gen_globalhk and gen_ml has started
	PostMessage(plugin.hwndParent, init_ipc, 0, 0);

	return 0;
}

STDAPI DllRegisterServer(void)
{
	// registers object, typelib and all interfaces in typelib
	HRESULT hr = _AtlModule.RegisterServer(true, &CLSID_Application);
	RegisterAdditionalResource(IDR_APPLICATION);
	return hr;
}

extern "C" {
	__declspec(dllexport) winampGeneralPurposePlugin * winampGetGeneralPurposePlugin()
	{
		return &plugin;
	}
	
}