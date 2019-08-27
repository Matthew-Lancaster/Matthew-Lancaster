
; ---------------------------------------------------------------
; I MADE MENU ITEM INTO INCLUDE FILE IN 3 PART 
; 01. INTRO SETUP MENU
; 02. THE MENU ROUTINE
; 03. ANY ROUTINE THE MENU USE
; ---------------------------------------------------------------
; SAVER OF RSI INJURY AND MORE ACCURATE
; THE INCLUDE FILE ARE SAME FOLDER
; ---------------------------------------------------------------
; FROM __ Sun 09-Jun-2019 07:03:00 __ Clipboard Count = 024
; TO   __ Sun 09-Jun-2019 10:00:00 __ Clipboard Count = 139 __ NEAR 3 HOUR
; AND THEN 
; TO   __ Sun 09-Jun-2019 17:48:00 __ Clipboard Count = 452 __ 10 HOURING 45 MINUTE
; ---------------------------------------------------------------

; Create the popup menu by adding some items to it.
if A_ScriptName=Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
{
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, RELAUNCH CODE, MenuHandler  ; Creates a new menu item.
}

if A_ScriptName=Autokey -- 21-AUTORUN.ahk
{
VAR_RUN_ME_NOW_AUTOBOOT=FALSE
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, RUN HERE NOW, MenuHandler  ; Creates a new menu item.
}


Menu, Tray, Add  
Menu, Tray, Add, Terminate Script, MenuHandler
Menu, Tray, Add, Terminate All AutoHotKey.exe -- LEFT(Ctrl)+WInKEY+ESCAPE, MenuHandler
Menu, Tray, Add  
Menu, Tray, Add, RESTORE_VB_KEEP_RUNNER AND ELITESPY -- RIGHT(Ctrl)+F1, MenuHandler
Menu, Tray, Add  
Menu, Tray, Add, RELOAD    ALL NET - VB CODE.exe, MenuHandler 
Menu, Tray, Add, TERMINATE ALL NET - VB CODE.exe, MenuHandler 
Menu, Tray, Add  
Menu, Tray, Add, RELOAD    All AutoHotKey Network.exe, MenuHandler  
Menu, Tray, Add, TERMINATE All AutoHotKey Network.exe, MenuHandler  
Menu, Tray, Add  

if A_ScriptName=Autokey -- 72-RUN HUBIC WITH DELAY.ahk
{
Menu, Tray, Add  
Menu, Tray, Add, 1 HOUR TIMER UNTIL HUBIC.EXE LEARN, MenuHandler  
}


if A_ScriptName=Autokey -- 58-Auto Repeat Browser Function Set.ahk
{
Menu, Tray, Add
Menu, Tray, Add, Pause __ Debby Hall, MenuHandler 
}

Menu, Tray, Add  
