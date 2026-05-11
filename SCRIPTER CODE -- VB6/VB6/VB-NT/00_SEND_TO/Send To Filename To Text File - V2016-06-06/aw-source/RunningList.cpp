/************************************************************************
RunningList.cpp

Implementation of running script dialog. Could use a better interface.

************************************************************************/

#include "stdafx.h"
#include "RunningList.h"
#include ".\runninglist.h"
#include "ScriptManager.h"

extern CScriptManager glbScriptManager;

LRESULT CRunningList::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	CAxDialogImpl<CRunningList>::OnInitDialog(uMsg, wParam, lParam, bHandled);
	RefreshList();
	return 1;  // Let the system set the focus
}
// CRunningList

void CRunningList::RefreshList(void)
{
	HWND hwndList = GetDlgItem(IDC_LIST1); 
	SendMessage(hwndList, LB_RESETCONTENT, 0, 0); 
	LISTSCRIPT::iterator si;
	int i = 0;
	for (si = runningScripts.begin(); si != runningScripts.end(); ++si)
	{
		SendMessage(hwndList, LB_ADDSTRING, 0, (LPARAM) ((CMyScriptSite*)*si)->m_desc.c_str()); 
		SendMessage(hwndList, LB_SETITEMDATA, i, (LPARAM) ((CMyScriptSite*)*si));
		i++;
	}
	::SetFocus(hwndList); 
}

LRESULT CRunningList::OnBnClickedButton1(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	RefreshList();
	return 0;
}

LRESULT CRunningList::OnBnClickedButton2(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	TCHAR szFileName[MAX_PATH]={0,};
	OPENFILENAME ofn;
	ZeroMemory(&ofn, sizeof(OPENFILENAME));
	ofn.lStructSize = sizeof(OPENFILENAME);
	ofn.hwndOwner = m_hWnd;
	ofn.lpstrTitle = "Load script";
	ofn.nMaxFile = MAX_PATH;
	ofn.lpstrFile = szFileName;
	ofn.lpstrFilter = "Script files (*.vbs)\0*.vbs\0"
		"Text files (*.txt)\0*.txt\0"
		"All files\0*.*\0";
	ofn.lpstrDefExt = "vbs";
	ofn.Flags = OFN_EXPLORER|OFN_FILEMUSTEXIST;
	char args[MAX_PATH];
	GetDlgItemText(IDC_ARGS, args, MAX_PATH);
	if (GetOpenFileName(&ofn)) {
			glbScriptManager.runfile(szFileName, false, args);
	}
	RefreshList();
	return 0;
}

LRESULT CRunningList::OnBnClickedButton3(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	glbScriptManager.stopall();
	RefreshList();
	return 0;
}

LRESULT CRunningList::OnBnClickedButton4(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	PostMessage(WM_CLOSE, 0, 0);
	return 0;
}
