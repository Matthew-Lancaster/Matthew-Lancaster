#pragma once

#include "Application.h"
#include "SiteManager.h"
#include "GEN.H"
#include "wa_ipc.h"
#include "wa_hotkeys.h"

extern long quit_ipc;
extern long hotkey_ipc;
extern long init_ipc;

extern winampGeneralPurposePlugin plugin;
extern winampMediaLibraryPlugin mlplugin;

extern WNDPROC lpWndProcOld;
extern WNDPROC lpPlWndProcOld;
extern HINSTANCE mhin;

extern const GUID CLSID_VBScript;

class CScriptManager
{
public:
	CScriptManager(void);
	~CScriptManager(void);

	char iniDir[MAX_PATH];
	DWORD dwr;
	DWORD dwr2;
	//list of runners
	void init();
	void run();
	void runfile(const char *fn, bool fromDir, const char *args);
	void stopall();
	int init1(void);
};
