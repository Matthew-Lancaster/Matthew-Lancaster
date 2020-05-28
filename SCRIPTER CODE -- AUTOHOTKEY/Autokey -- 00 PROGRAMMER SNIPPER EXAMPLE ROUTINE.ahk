#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
; -------------------------------------------------------------------
#SingleInstance force
; -------------------------------------------------------------------
; #Persistent
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

DetectHiddenWindows, oFF
SetTitleMatchMode 3  ; Specify Full path

; -------------------------------------------------------------------
; #InstallKeybdHook
; #InstallMouseHook
; -------------------------------------------------------------------


LOOP
{				                    ; Place cursor over [location]
	MouseGetPos, offsetx, offsety	; x x x
	offsetx := offsetx - 0	     	; x O x  O = tip of mouse cursor
	offsety := offsety - 0	     	;
	WinGetTitle, ACTIVE_TITLE, A
	GetKeyState, state, LButton 

	IfWinActive Calculator ahk_class CalcFrame
	IF offsetx>41
	IF offsetx<562
	IF offsety<38   ; --  HITT IN TITLE BAR AREA NOT ANY OTHER BUTTON -- DEPEND SIZE BUTTON TYPE THING
	IF (state = "D")
	{
		; ---------------------------------------------------------------
		; ORIGINAL INTENTION CLICK TITLE BAR CLOSE APP LOAD ANOTHER
		; CALC FOR WINDOWS 7 ONLY RUN BY SPECIAL APP 
		; REPLACE GET WINDOWS 10 VERSION 
		; ONLY BY TASK-BAR LINK 
		; AND HOTKEY NOT DETECTABLE
		; ---------------------------------------------------------------
		; ; TOOLTIP % "Offset x = " . offsetx . ", y = " . offsety
		WINCLOSE Calculator ahk_class CalcFrame
		RUN, "C:\PStart\# NOT INSTALL REQUIRED\CALC WIN 04 10\Calc Windows 10.exe"
	}

}
RETURN



; -------------------------------------------------------------------
; 0001 __ TEST ROUTINE
; -------------------------------------------------------------------
GLOBAL ArrayCount
; Each array must be initialized before use:
FN_Array_1 := []
FN_Array_2 := []
FN_Array_3 := []
DATE_MOD_Array := []

ArrayCount=0
ArrayCount+=1
FN_Array_1[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
FN_Array_2[ArrayCount]:="EliteSpy+ by Andrea"
ArrayCount+=1
FN_Array_1[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
FN_Array_2[ArrayCount]:="VB_KEEP_RUNNER"
ArrayCount+=1
FN_Array_1[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\CPU % OF A PROGRAM\CPU % INDIVIDUAL PROCESS.exe"
FN_Array_2[ArrayCount]:="INDIVIDUAL PROCESS"

ArrayCount+=1
FN_Array_1[ArrayCount]:=""
FN_Array_2[ArrayCount]:=""


Loop % ArrayCount
{
	Element := FN_Array_1[A_Index]
	IF !Element
	{
		ArrayCount-=1
	}
}
MSGBOX %ArrayCount%
EXITAPP
RETURN
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; 0002 __ MSGBOX
; -------------------------------------------------------------------
; 4096 BRING TO FOREGROUND
; 4100 BRING TO FOREGROUND + YES NOT
; FOREGROUND + YES NOT
; 4096 + 4
; -----------------------------------------------
; MSGBOX ,4100,,% "------ -- ------ ----- ---- ---- ---------`n`n"
; IFMSGBOX YES
; -----------------------------------------------
MSGBOX ,4096,,NOT GOODSYNC PORTABLE D-DRIVE WITH ASUS 4G AT THE MOMENT`n`nAS IMAGE AT GOOGLE PHOTO NOT IN-LINE YET
RETURN
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; #InstallKeybdHook
; #InstallMouseHook
; -------------------------------------------------------------------
; DETECTOR LAST KEYCODE AND 2ND SC CODE
; -------------------------------------------------------------------
; LOOP
; {
	; TOOLTIP % A_PRIORKEY
; }
; RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; ; -------------------------------------------------------------------
; ; -------------------------------------------------------------------
; SetFormat, Integer, Hex
; Gui +ToolWindow -SysMenu +AlwaysOnTop
; Gui, Font, s14 Bold, Arial
; Gui, Add, Text, w100 h33 vSC 0x201 +Border, {SC000}
; Gui, Show,, % "// ScanCode //////////"
; LOOP
; {
; Loop 9
  ; OnMessage( 255+A_Index, "ScanCode" ) ; 0x100 to 0x108
; }

; ScanCode( wParam, lParam ) {
 ; Clipboard_TMP := "SC" SubStr((((lParam>>16) & 0xFF)+0xF000),-2) 
 ; GuiControl,, SC, %Clipboard_TMP%
; }
; RETURN
; ; -------------------------------------------------------------------


