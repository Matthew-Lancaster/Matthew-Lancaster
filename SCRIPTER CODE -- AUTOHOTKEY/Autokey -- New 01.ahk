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
FN_Array_1[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB KEEP RUNNER.exe"
FN_Array_2[ArrayCount]:="VB KEEP RUNNER"
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
