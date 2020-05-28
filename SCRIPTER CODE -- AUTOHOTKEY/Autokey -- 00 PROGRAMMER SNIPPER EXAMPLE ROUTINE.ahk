#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
; -------------------------------------------------------------------
#SingleInstance force
; -------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
SetStoreCapslockMode, off

DetectHiddenWindows, oFF
SetTitleMatchMode 3  ; Specify Full path


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


