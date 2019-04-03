;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 02-SAVE AS KEY ENTER.ahk
;# __ 
;# __ Autokey -- 02-SAVE AS KEY ENTER.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================


;# Thu 25 Aug 2016 01:47:51 Work Time Beginner
;# Thu 25 Aug 2016 03:00:00 3 Hour
;# Thu 25-Aug-2016 06:16:41 2nd Wave of Work Better Coder Emerged
;# Thu 25-Aug-2016 16:31:00 40 Minute
;# Thu 25-Aug-2016 11:20:00 3RD WAVE 
;# Thu 25-Aug-2016 16:31:00 5 & Half Hour
;# Sat 21-Apr-2018 07:24:30 4TH WAVE Massively Better
;# Sat 21-Apr-2018 08:28:00 1 Hour
;# __ 

;--------------------------------------------------------------------
; AutoHotKeys Project Code to User Press Enter Key on (Save As) Window  Dialogue in Chrome.exe & Exit On a Timer 120 Sec When Chrome.exe Gone

;--------------------------------------------------------------------
; 001 First Wave Bash
; -------------------------------------------------------------------
; Thu 25 Aug 2016 01:47:51 -- Length Time 1hr 45M - Finish Code
; Thu 25 Aug 2016 02:35:00 -- Length Time+2hr 30M - Fine Tune & Doc
; Thu 25 Aug 2016 03:00:00 -- Length Time+3hr 00M - Publish On-line
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 002 A Second Session Bash __ Dawn
; -------------------------------------------------------------------
; Thu 25-Aug-2016 06:16:41 --
; Thu 25 Aug 2016 07:00:00 -- Length Time 0 hr 40 Min -- Tune Doc & Publish
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 003 A Third Session Bash __ Morning
; -------------------------------------------------------------------
; Thu 25-Aug-2016 11:20:00 
; Thu 25-Aug-2016 14:59:29
; Thu 25-Aug-2016 16:31:00 -- 11 TO 4pm 5 Half Hour Hurt Ouch
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 004 A Fourth Session Improved Code Dramatically
; -------------------------------------------------------------------
; Sat 21-Apr-2018 07:24:30 -- BEGINNER
; Sat 21-Apr-2018 08:28:00 -- 1 HOUR WITH PUBLISH ON-LINE
;--------------------------------------------------------------------

; -------------------------------------------------------------------
; 005 ANOTHER BASH CPU USAGE WAS A BIT HIGHER THAN WANTED EVEN ON PAUSE 
; -------------------------------------------------------------------
; Thu 05-Jul-2018 09:44:14
; Thu 05-Jul-2018 09:48:00
; -------------------------------------------------------------------

;--------------------------------------------------------------------
; 2nd Project in a Line of AutoKey -- First On Blogger Here
; This is All Copy Paste in and Go Project for AutoHotKeys Working
; Building Block Head Start
;--------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location Internet
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set Google Drive
; Possible Censorship of Code Detected By Google as Malicious Happen Here
; unlike DropBox that has All Available
; https://drive.google.com/open?id=0BwoB_cPOibCPTnRZZVFuRFpHOTg
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set DropBox
; https://www.dropbox.com/sh/ntghoncyb8py1tf/AACWYrfkVn9PlqpYzNNSMcpMa?dl=0
;--------------------------------------------------------------------
; Link to This File On DropBox With Most Up to Date
; https://www.dropbox.com/s/80o59c02zjjta90/Autokey%20--%2002-SAVE%20AS%20KEY%20ENTER.ahk?dl=0
;# ------------------------------------------------------------------

; BEGIN CODE
; -------------------------------------------------------------------
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance force
; -------------------------------------------------------------------
; SingleInstance Force Reduce Question at Reload
; -------------------------------------------------------------------

;--------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

; Create the popup menu by adding some items to it.
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, Terminate Script, MenuHandler  ; Creates a new menu item.
Menu, Tray, Add, Terminate All AutoHotKey.exe, MenuHandler  ; Creates a new menu item.

SoundBeep , 2000 , 200
SoundBeep , 2500 , 100
SetStoreCapslockMode, off

Var_Count=0

pause
soundbeep , 1100,100

SetTimer TIMER_Main_Code,1000
SetTimer TIMER_PREVIOUS_INSTANCE,1

Return

;----------------------------------------------------------------------
TIMER_Main_Code:

	If WinExist("Save As")
	{

		If WinExist("ahk_class Chrome_WidgetWin_1")
		{
			WinGet, HWND, ID, Save As
			WinGetClass, This_Class, ahk_id %HWND%
			WinGet, path, ProcessName, ahk_id %HWND%
			WinGetText, OutputVar, ahk_id %HWND%
			WinGetPos, WinLeft, WinTop, WinWidth, WinHeight, ahk_id %HWND%

			; Multiple If-And Statement With Separated Lines
			; ----------------------------------------------
			IF (This_Class="#32770" 
			and PATH="Chrome.exe")
			{
				SoundBeep , 1200, 100
				SoundBeep , 2500 , 100
				ControlClick, &Save, Save As
			}
		}
		
	}

	WinGet, HWND, ID, A
	WinGetClass, This_Class, ahk_id %HWND%
	IF This_Class=Chrome_WidgetWin_1 
		Var_Count = 0

	IF This_Class <> Chrome_WidgetWin_1 
		Var_Count += 1
		
	if Var_Count > 120
	{
		SoundBeep , 1100,100
		Pause
		SoundBeep , 1100,100
		Var_Count = 0
	}
	
Return


MenuHandler:
	; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	if A_ThisMenuItem=Terminate Script
		Process, Close,% DllCall("GetCurrentProcessId")
	
	if A_ThisMenuItem=Terminate All AutoHotKey.exe
	{
		Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS" /F /IM AutoHotKey.exe /T , , Max
		
		;  ----------------------------------------------------------
		; PROBLEM HERE IF PROGRAM THAT CALL THE BATCH FILE IS KILL SO IS THEN BATCH FILE
		; AND WE GET OVER THAT BY GO EXTRA VIA VBSCRIPT ANOTHER FILE
		; COULD OF RUN A  LOOP AND KILL BUT TRY NOT LOSE OWN ONE FIRST
		; [ Saturday 14:55:00 Pm_02 March 2019 ]
		;  ----------------------------------------------------------

		;  ----------------------------------------------------------
		; OTHER OPTION SET PROCESS KILLER
		;  ----------------------------------------------------------
		; Run, BAT_03_PROCESS_KILLER.BAT /F /IM AutoHotKey.exe /T , , Max
		; Run, %ComSpec% /k ""BAT_03_PROCESS_KILLER.BAT" "/F" "/IM" "AutoHotKey.exe" "/T"" , , Max
		; Process, Close, AutoHotKey.exe
		;  ----------------------------------------------------------
	
		; AUTO GENERATED FILE BY HERE VISUAL BASIC ORIGINAL LONG BEFORE AUTOHOTKEY WANT
		; D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB KEEP RUNNER.exe
		; D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe
		; -------------------------------------------------------------------
		; AND USED BY HERE
		; LOT OF AUTOHOTKEYS TRAY MENU ITEM
		; -------------------------------------------------------------------
		; [ Saturday 14:52:10 Pm_02 March 2019 ]
		; -------------------------------------------------------------------
		; EDITOR COPY PASTE FROM VBS 39-KILL PROCESS.VBS
		; THIS FILE BECAME USE BY
		; LOT OF AUTOHOTKEYS TRAY MENU ITEM
		; AND THEY USE IT HERE THIS ONE
		; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT
		; ORIGINAL AT HERE LOCATION 
		; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
		; AND MOVED HERE MAYBE 
		; -------------------------------------------------------------------
		; MOST LIKELY TRY AND KEEP IN SYNC LATER
		; EXCEPT THE AUTO GENERATOR
		; -------------------------------------------------------------------
}
return



TIMER_PREVIOUS_INSTANCE:
SetTimer TIMER_PREVIOUS_INSTANCE,10000
if ScriptInstanceExist()
{
    Exitapp
}
return

ScriptInstanceExist() {
    static title := " - AutoHotkey v" A_AhkVersion
    dhw := A_DetectHiddenWindows
    DetectHiddenWindows, On
    WinGet, match, List, % A_ScriptFullPath . title
    DetectHiddenWindows, % dhw
    return (match > 1)
    }
Return

;# ------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ------------------------------------------------------------------


;# ------------------------------------------------------------------
ExitFunc(ExitReason, ExitCode)
{
    if ExitReason not in Logoff,Shutdown
    {
        ;MsgBox, 4, , Are you sure you want to exit?
        ;IfMsgBox, No
        ;    return 1  ; OnExit functions must return non-zero to prevent exit.
    }
    ; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}

class MyObject
{
    Exiting()
    {
        ;
        ;MsgBox, MyObject is cleaning up prior to exiting...
        /*
        this.SayGoodbye()
        this.CloseNetworkConnections()
        */
    }
}
;# ------------------------------------------------------------------
; exit the app


;--------------------------------------------------------------------
; Work Note Rather Novice at First Began
;--------------------------------------------------------------------
; Yes Fixed the Error Variable Increment -- Seem Like a C++ Method Style -- TEST_VAR += 1

; This Program Has Problem in Win 10 The Dialog to Save Do Stay So Move On After Save to Next Icon to Do Save JPG Screenshot
; ----
; Save as Shortcut - Chrome Web Store
; https://chrome.google.com/webstore/detail/save-as-shortcut/flehofiklehmnnolpjcamplcnmhgcbkk?utm_source=chrome-app-launcher-info-dialog
; ----

; A Tip to Learn With This Program is 
; When Ask to Save Select JPG From PNG and It will Remember
; ----
; Capture Webpage Screenshot Entirely. FireShot - Chrome Web Store
; https://chrome.google.com/webstore/detail/capture-webpage-screensho/mcbpblocgmgfnpjjppndjkmgjaogfceg?utm_source=chrome-app-launcher-info-dialog
; ----
;--------------------------------------------------------------------
; Work Logg -- Start time not that clear
; Yes Found Time Begin Phew Charging Someone for This Automation -- RSI Injury


;--------------------------------------------------------------------
;--------------------------------------------------------------------
; Info Note
;--------------------------------------------------------------------

; Use of these 4 Link in Turn to Save Page of the Chrome Extension
;--------------------------------------------------------------------
; 1.. Copy All Urls - Chrome Web Store
; 2.. Save as Shortcut - Chrome Web Store
; 3.. Capture Webpage Screenshot Entirely. FireShot - Chrome Web Store
; 4.. SingleFile - Chrome Web Store -- to Save an HTML Version
;--------------------------------------------------------------------

;--------------------------------------------------------------------
;--------------------------------------------------------------------
;Ref: - Credit to Source Helper Ref:
;--------------------------------------------------------------------
;--------------------------------------------------------------------
;AutoKey Source to Learn
;--------------------------------------------------------------------
;Begin Here Google Search
;--------------------------------------------------------------------
;1..
;How to save files using Enter key? - Ask for Help - AutoHotkey Community
;https://autohotkey.com/board/topic/124006-how-to-save-files-using-enter-key/

;--------------------------------------------------------------------
;AutoKey Info Help Web
;--------------------------------------------------------------------
;1..
;OnExit
;https://autohotkey.com/docs/commands/OnExit.htm
;--------
;2..
;While-loop
;https://autohotkey.com/docs/commands/While.htm
;--------
;3..
;Else
;https://autohotkey.com/docs/commands/Else.htm
;--------
;4..
;OnExit
;https://autohotkey.com/docs/commands/OnExit.htm
;--------
;5..
;Functions
;https://autohotkey.com/docs/Functions.htm
;--------------------------------------------------------------------

;--------------------------------------------------------------------
;Ref
;Chrome Extension
;--------------------------------------------------------------------

;1..
;--------
;This One not Much To do With Here But Still Used
;--------
;Copy All Urls - Chrome Web Store
;https://chrome.google.com/webstore/detail/copy-all-urls/djdmadneanknadilpjiknlnanaolmbfk?utm_source=chrome-app-launcher-info-dialog
;--------

;2..
;--------
;Yes Save to URL to Download Folder
;--------
;Save as Shortcut - Chrome Web Store
;https://chrome.google.com/webstore/detail/save-as-shortcut/flehofiklehmnnolpjcamplcnmhgcbkk?utm_source=chrome-app-launcher-info-dialog
;--------

;3..
;--------
;This One Save Auto to Download Folder Without Our Coding Requirement
;--------
;Capture Webpage Screenshot Entirely. FireShot - Chrome Web Store
;https://chrome.google.com/webstore/detail/capture-webpage-screensho/mcbpblocgmgfnpjjppndjkmgjaogfceg?utm_source=chrome-app-launcher-info-dialog
;--------

;4..
;--------
;As Well as Save Screen Shot About Do a Web Page HTML Save also
;Be Good If Done Without an Extra Ask to Download Save As
;But Setup to Store a Page on a Web On-line Auto -- PageArchiver
;--------
;SingleFile - Chrome Web Store
;https://chrome.google.com/webstore/detail/singlefile/mpiodijhokgodhhofbcjdecpffjipkle?utm_source=chrome-app-launcher-info-dialog
;--------------------------------------------------------------------


;--------------------------------------------------------------------
;SEMI USELESS CODE SHOVED TO BOTTOM LOOK AT LATER
;--------------------------------------------------------------------
;WinGet nChromeWindows, Count, ahk_class Chrome_WidgetWin_1

;----
;[Solved] How to get the hwnd of the current script - Ask for Help - AutoHotkey Community
;https://autohotkey.com/board/topic/35140-solved-how-to-get-the-hwnd-of-the-current-script/
;----
;ThisScriptsHWND := WinExist("Ahk_PID " DllCall("GetCurrentProcessId"))
;msgbox %ThisScriptsHWND% is this script's HWND

;--------------------------------------------------------------------
;Don't stay RSI -Save Me- if Form Change
;--------------------------------------------------------------------
;WinGetTitle, active_title, A
;IF active_title <> O_active_title
;{
;	SoundBeep
;	break
;}
;O_active_title = %active_title%