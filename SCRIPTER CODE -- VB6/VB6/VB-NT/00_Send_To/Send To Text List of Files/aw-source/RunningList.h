// RunningList.h : Declaration of the CRunningList

#pragma once

#include "resource.h"       // main symbols
#include <atlhost.h>

#include "myscriptsite.h"
#include "SiteManager.h"
#include <list>

typedef std::list <CMyScriptSite*> LISTSCRIPT;
extern LISTSCRIPT runningScripts;
extern long quit_ipc;
// CRunningList
class CRunningList : 
	public CAxDialogImpl<CRunningList>
{
public:
	CRunningList()
	{
	}

	~CRunningList()
	{
	}

	enum { IDD = IDD_RUNNING };

BEGIN_MSG_MAP(CRunningList)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_CLOSE, OnClose)
	COMMAND_ID_HANDLER(IDC_LIST1, ListBHandle)
	COMMAND_HANDLER(IDC_BUTTON1, BN_CLICKED, OnBnClickedButton1)
	COMMAND_HANDLER(IDC_BUTTON2, BN_CLICKED, OnBnClickedButton2)
	COMMAND_HANDLER(IDC_BUTTON3, BN_CLICKED, OnBnClickedButton3)
	COMMAND_HANDLER(IDC_BUTTON4, BN_CLICKED, OnBnClickedButton4)
	CHAIN_MSG_MAP(CAxDialogImpl<CRunningList>)
END_MSG_MAP()


LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	//return CAxDialogImpl<CRunningList>::EndDialog(0);
	return CAxDialogImpl<CRunningList>::DestroyWindow();
}

LRESULT ListBHandle(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	if (wNotifyCode == LBN_SELCHANGE)
	{
		HWND hwndList = GetDlgItem(IDC_LIST1); 		
		int nItem = SendMessage(hwndList, LB_GETCURSEL, 0, 0); 
		CMyScriptSite* itm = (CMyScriptSite*)SendMessage(hwndList, LB_GETITEMDATA, nItem, 0); 

		SendMessage(plugin.hwndParent, quit_ipc, (WPARAM)itm, (LPARAM)0);
		RefreshList();
	}
	return 0;
}

void RefreshList(void);
LRESULT OnBnClickedButton1(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
LRESULT OnBnClickedButton2(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
LRESULT OnBnClickedButton3(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
LRESULT OnBnClickedButton4(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
};