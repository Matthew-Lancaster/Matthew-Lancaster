tray_iconCleanup()





tray_icons()
{
static TB_BUTTONCOUNT := 0x0418, TB_GETBUTTON := 0x0417
array := []
array[0] := ["sProcess","toolTip","nMsg","uID","idx","idn","Pid","hwnd","sClass","hIcon"]
Index := 0
detectHiddenWindows_old := a_detectHiddenWindows
detectHiddenWindows, on
trayWindows := "Shell_TrayWnd,NotifyIconOverflowWindow"
loop, parse, trayWindows, csv
{
winGet, taskbar_pid, PID, ahk_class %a_loopField%
hProc := dllCall( "OpenProcess", "Uint", 0x38, "int", 0, "Uint", taskbar_pid )
pProc := dllCall( "VirtualAllocEx", "Uint", hProc, "Uint", 0, "Uint", 32, "Uint", 0x1000, "Uint", 0x4 )
idxTB := tray_getTrayBar()
sendMessage, %TB_BUTTONCOUNT%, 0, 0, ToolbarWindow32%idxTB%, ahk_class %a_loopField%
loop, %errorLevel%
{
sendMessage, %TB_GETBUTTON%, a_index-1, pProc, ToolbarWindow32%idxTB%, ahk_class %a_loopField%
varSetCapacity( btn,32,0 ), varSetCapacity( nfo,32,0 )
dllCall( "ReadProcessMemory", "Uint", hProc, "Uint", pProc, "Uint", &btn, "Uint", 32, "Uint", 0 )
iBitmap := numGet( btn, 0 )
idn := numGet( btn, 4 )
Statyle := numGet( btn, 8 )
if dwData := numGet( btn,12,"Uint" )
iString := numGet( btn,16 )
else dwData := numGet( btn,16,"int64" ), iString := numGet( btn,24,"int64" )
dllCall( "ReadProcessMemory", "Uint", hProc, "Uint", dwData, "Uint", &nfo, "Uint", 32, "Uint", 0 )
if numGet( btn,12,"Uint" )
{
hwnd := numGet( nfo, 0 )
uID := numGet( nfo, 4 )
nMsg := numGet( nfo, 8 )
hIcon := numGet( nfo,20 )
}
else hwnd := numGet( nfo, 0,"int64" ), uID := numGet( nfo, 8,"Uint" ), nMsg := numGet( nfo,12,"Uint" )
winGet, pid, PID, ahk_id %hwnd%
winGet, sProcess, ProcessName, ahk_id %hwnd%
winGetClass, sClass, ahk_id %hwnd%
 
varSetCapacity( sTooltip,128 ), varSetCapacity( wTooltip,128*2 )
dllCall( "ReadProcessMemory", "Uint", hProc, "Uint", iString, "Uint", &wTooltip, "Uint", 128*2, "Uint", 0 )
dllCall( "WideCharToMultiByte", "Uint", 0, "Uint", 0, "str", wTooltip, "int", -1, "str", sTooltip, "int", 128, "Uint", 0, "Uint", 0 )
idx := a_index-1
toolTip := a_isUnicode ? wTooltip : sTooltip
Index++
for a,b in array[0]
array[Index,b] := %b%
}
dllCall( "VirtualFreeEx", "Uint", hProc, "Uint", pProc, "Uint", 0, "Uint", 0x8000 )
dllCall( "CloseHandle", "Uint", hProc )
}
detectHiddenWindows, %detectHiddenWindows_old%
return array
}
 
tray_iconCleanup()
{
/*
remove orphan icons ( Their processes have been killed )
*/
tray_icons := tray_icons()
for index in tray_icons
{
ifEqual,index,0,continue
if ( tray_icons[index, "sProcess"] = "" )
tray_iconRemove( tray_icons[index, "hwnd"], tray_icons[index, "uID"],"", tray_icons[index, "hIcon"] )
}
}
 
tray_iconRemove( hwnd, uID, nMsg = 0, hIcon = 0, nRemove = 0x2 )
{
/*
NIM_DELETE := 0x00000002
 
typedef struct _NOTIFYICONDATA {
DWORD cbSize; x00
hwnd hwnd; x04
UINT uID; x0c/d12
UINT uFlags; x10/d16
UINT uCallbackMessage; x14/d20
HICON hIcon; x18/d24
TCHAR szTip[64]; x20 ( UNICODE )
DWORD dwState; xa0
DWORD dwStateMask; xa4
TCHAR szInfo[256]; xa8 ( UNICODE )
union {
UINT uTimeout; x01a8
UINT uVersion; x01a8
};
TCHAR szInfoTitle[64]; x01ac ( UNICODE )
DWORD dwInfoFlags; x022c
GUID guidItem; ( 128bits ) x0230
HICON hBalloonIcon; x023c
;end x0244
} NOTIFYICONDATA, *PNOTIFYICONDATA;
 
MSDN:
http://msdn.microsoft.com/en-us/library/windows/desktop/bb773352%28v=vs.85%29.aspx
*/
varSetCapacity( nid,size := 936+4*a_ptrSize )
numPut( size, nid, 0, "uint" )
numPut( hwnd, nid, a_ptrSize )
numPut( uID, nid,a_ptrSize*2, "uint" )
numPut( 1|2|4, nid,a_ptrSize*3, "uint" )
numPut( nMsg , nid,a_ptrSize*4, "uint" )
numPut( hIcon, nid,a_ptrSize*5, "uint" )
return dllCall( "shell32\Shell_NotifyIconA", "Uint", nRemove, "Uint", &nid )
}
 
tray_iconHide( idn, bHide = true )
{
static TB_HIDEBUTTON := 0x404, WM_SETTINGCHANGE := 0x001A
idxTB := tray_getTrayBar()
sendMessage, %TB_HIDEBUTTON%, idn, bHide, ToolbarWindow32%idxTB%, ahk_class Shell_TrayWnd	;NotifyIconOverflowWindow
sendMessage, %WM_SETTINGCHANGE%, 0, 0, , ahk_class Shell_TrayWnd
}
 
tray_iconDelete( idx )
{
static TB_DELETEBUTTON := 0x416, WM_SETTINGCHANGE := 0x001A
idxTB := tray_getTrayBar()
sendMessage, %TB_DELETEBUTTON%, idx - 1, 0, ToolbarWindow32%idxTB%, ahk_class Shell_TrayWnd
sendMessage, %WM_SETTINGCHANGE%, 0, 0, , ahk_class Shell_TrayWnd
}
 
tray_iconMove( idxOld, idxNew )
{
static TB_MOVEBUTTON := 0x452
idxTB := tray_getTrayBar()
sendMessage, %TB_MOVEBUTTON%, idxOld - 1, idxNew - 1, ToolbarWindow32%idxTB%, ahk_class Shell_TrayWnd
}
 
tray_getTrayBar()
{
detectHiddenWindows_old := a_detectHiddenWindows
detectHiddenWindows, on
controlGet, hParent, hwnd,, TrayNotifyWnd1 , ahk_class Shell_TrayWnd
controlGet, hChild , hwnd,, ToolbarWindow321, ahk_id %hParent%
loop
{
controlGet, hwnd, hwnd,, ToolbarWindow32%a_index%, ahk_class Shell_TrayWnd
if !hwnd
break
else if ( hwnd = hChild )
{
idxTB := a_index
break
}
}
detectHiddenWindows, %detectHiddenWindows_old%
return idxTB
}
