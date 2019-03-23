#pragma once
#include "Gen.h"
#include "wa_ipc.h"
#include "ml.h"
#include "application.h"

WNDPROC lpWndProcOld;
WNDPROC lpPlWndProcOld;
WNDPROC lpMainWndProcOld;

void config();
void quit();
void quitML();
int init();
int PluginMessageProc(int message_type, int param1, int param2, int param3);
LRESULT CALLBACK PlWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK MWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
int MLWndProc(int message, int param1, int param2, int param3);

HINSTANCE mhin;
HWND HwndPl;
//HANDLE quitSignal;
long quit_ipc;
long hotkey_ipc;
long init_ipc;
DWORD pappCookie;

winampGeneralPurposePlugin plugin =
{
	GPPHDR_VER,
		"ActiveWinamp - v1.0",
		init,
		config,
		quit,
};

winampMediaLibraryPlugin mlplugin =
{
	MLHDR_VER,
		"ActiveWA support",
		init,
		quitML,
		MLWndProc,
};