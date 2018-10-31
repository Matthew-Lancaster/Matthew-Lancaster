;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 43-Test If File Change Iin Folder.ahk
;# __ 
;# __ Autokey -- 43-Test If File Change Iin Folder.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;# __ DATE BEGIN
;# __ Thu 25-Oct-2018 18:51:55 CODE KNOCK UP TIME
;# __ Thu 25-Oct-2018 19:41:00
;# __ 
;  =============================================================

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

#Persistent
;IT USER ExitFunc TO EXIT FROM #Persistent
;OR      Exitapp  TO EXIT FROM #Persistent
;Exitapp CALLS ONTO ExitFunc
;--------------------
;--------------------
#SingleInstance force
;--------------------

;# ------------------------------------------------------------------
; ---- LOCATION ONLINE
; -------------------------------------------------------------------
; ----
; ----
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; DESCRIPTION
;# ------------------------------------------------------------------
; TEST HUBIC HOW LONG TAKE TO BEGIN RESTORE AFTER ISSUE COMMAND
; TAKE LONGER TIME
; NOT SURE IF LONGER THE BIGGER THE RESTORE WORK
; 2-TB HARD DRIVE CRASH AND REQUIRE RESTORER D DRIVE
;# ------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; ----
; -------------------------------------------------------------------
; FROM   Thu 25-Oct-2018 18:51:55 CODE KNOCK UP TIME
; TO     Thu 25-Oct-2018 19:41:00
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; GUI CREDIT SOURCE
; -------------------------------------------------------------------
; ----
; list of all processes - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/118907-list-of-all-processes/
; ----
;# ------------------------------------------------------------------

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------
SoundBeep , 1500 , 400
; -------------------------------------------------------------------

; GLOBAL SETTINGS ===================================================

GLOBAL RUN_COUNT
GLOBAL RUN_1
RUN_1 = FALSE
RUN_COUNT := 1 

	; Each array must be initialized before use:
;Name
global FN_Array := []
global STATUS_DONE_Array := []
global STATUS_BEGIN_Array := []
global STATUS_END_IT_Array := []
global STATUS_DESCRIPTION_Array := []

GLOBAL SET_SHOW
GLOBAL O_SET_SHOW_2
GLOBAL O_MATTER


; DUPLICATE ENTRY REM-ED OUT - TURNED OFF
; -------------------------------------------------------------------
; #Warn
; #NoEnv
; #SingleInstance Force

; ADDITIONAL CODE PART CANNOT SEEM TO USE A GLOBAL VARIABLE THROUGH A FUNCTION PROBLEM AT THE MOMENT
; MAYBE HERE 
; ----
; command line - Autohotkey / AHK - Param %1% not accessible within a function - Stack Overflow
; https://stackoverflow.com/questions/49905729/autohotkey-ahk-param-1-not-accessible-within-a-function
; ----
;
; GLOBAL REASON_PROCESS_WAIT
; REASON_PROCESS_WAIT=
; -------------------------------------------------------------------

GOSUB GUI_INIT_1
GOSUB GUI_INIT_2

GOSUB TEST_FILE_IN_FOLDER
 
Settimer, TEST_FILE_IN_FOLDER,1000

RETURN


GUI_INIT_1:

	; GUI ===============================================================
	Gui, 1: Margin, 5, 5
	gui, 1: font, s14 ; , Arial ; , Calibri  

	; Gui, Add, Edit, xm ym w200 hWndhSearch vprcsearch
	; DllCall("user32.dll\SendMessage", "Ptr", hSearch, "UInt", 0x1501, "Ptr", 1, "Str", "Prozess Name here", "Ptr")

	width_x=1150

	Gui, 1: Add, Button, y+5 w%width_x% gSTART_RERUN, Test If File or Folder Found in Folder
	Gui, 1: add, Button, y+5 W%width_x% vStatus, Status

	; call it a v vStatus rather than a GStatus and then you are able to update button

	Gui, 1: Add, ListView, xm y+5 w%width_x% h240, Time_Begin|Time_Ender|Name|HardCore
	LV_ModifyCol(1, 300)
	LV_ModifyCol(2, 300)
	LV_ModifyCol(3, 150)
	LV_ModifyCol(4, 370)
	Menu, Tray, Add  ; Creates a separator line.
	Menu, Tray, Add, Restore Window, MenuHandler  ; Creates a new menu item.

	Gui, 1: Show, AutoSize,

RETURN

GUI_INIT_2:

	FN_Array := []
	STATUS_DONE_Array := []
	STATUS_BEGIN_Array := []
	STATUS_END_IT_Array := []
	STATUS_DESCRIPTION_Array := []

	SET_SHOW=
	O_MATTER=
	O_SET_SHOW_2=

	GuiControl, 1:, Status, Wait For File or Folder To Exist in Folder

RETURN
	

; SCRIPT ============================================================
START_RERUN:
	Soundbeep 1000,100
	GOSUB GUI_INIT_2
RETURN

Status:
RETURN

MenuHandler:
; MsgBox You selected %A_ThisMenuItem% from menu %A_ThisMenu%.

IF A_ThisMenuItem=Restore Window
	Gui, 1: Show


return


TEST_FILE_IN_FOLDER:

	ArrayCount := 0
	ArrayCount += 1
	FN_Array[ArrayCount]:="K:\4\" ;7G ; The folder here.
	ArrayCount += 1
	FN_Array[ArrayCount]:="D:\0\" ;8M 
	ArrayCount += 1
	FN_Array[ArrayCount]:="D:\1\" ;7G 
	ArrayCount += 1
	FN_Array[ArrayCount]:="I:\1\" ;4G
	ArrayCount += 1
	FN_Array[ArrayCount]:="I:\2\" ;4G

	Loop % ArrayCount
	{
		Folder := FN_Array[A_Index]

		FormatTime, TimeString, %a_now%, ddd MMM d, yyyy -- hh:mm:ss tt

		IF !STATUS_BEGIN_Array[A_Index]
		{
			STATUS_BEGIN_Array[A_Index]:=TimeString
			STATUS_DESCRIPTION_Array[A_Index]:="Wait For File or Folder to Exist in Folder"
		}
		
		; How to make a calculator
		; ----
		; complete addition and multiplication from clipboard - Page 2 - Ask for Help - AutoHotkey Community
		; https://autohotkey.com/board/topic/123667-complete-addition-and-multiplication-from-clipboard/page-2
		; ----
		X_Compare := (20*ArrayCount)

		IF (RUN_COUNT>X_Compare and RUN_COUNT>0)
		{
			RUN_COUNT:=-1
			Gui, 1: Hide
		}
		IF RUN_COUNT>0
			RUN_COUNT+=1

		SET_GO=TRUE
		IF (A_ComputerName = "1-ASUS-X5DIJ") 
		IF (A_ComputerName = "2-ASUS-EEE") 
		IF (A_ComputerName = "3-LINDA-PC") 
		IF (A_ComputerName = "5-ASUS-P2520LA") 
		IF (A_ComputerName = "7-ASUS-GL522VW") 
		SET_GO=FALSE

		
		Number=
		loop, %Folder%*.*
			Number := A_Index
		
		
		if !Number
		Loop, %Folder%*, 2
		{
			;folder := A_LoopFileName
			Number=1
			break
		}

		if Number
		{
			; Settimer TEST_FILE_IN_FOLDER,off

			IF !STATUS_END_IT_Array[A_Index]
				STATUS_END_IT_Array[A_Index]:=TimeString

			GuiControl, 1:, Status, FILE FOUND
			
			STATUS_DESCRIPTION_Array[A_Index]:="FILE OR FOLDER FOUND"

			MATTER=
			Loop % ArrayCount
			{
			VAR_MATTER:=STATUS_DESCRIPTION_Array[A_Index]
			MATTER=%MATTER%%VAR_MATTER%
			}
			
			IF MATTER<>%O_MATTER%
			{
				O_MATTER=%MATTER%
				; ---------------------------------------------------------------
				; Restore:
				; ---------------------------------------------------------------
				IF !SET_SHOW
				{
					Gui, 1: Show
					SET_SHOW:=1
				}
			}
		}

	}
	
	LV_Delete()
	Loop % ArrayCount
	{
	
		LV_Add("", STATUS_BEGIN_Array[A_Index], STATUS_END_IT_Array[A_Index], FN_Array[A_Index],STATUS_DESCRIPTION_Array[A_Index])
	}
	
	; DetectHiddenWindows, ON
		
	MY_AHK=C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 43-Test If File Change Iin Folder.ahk
	AHK_TERMINATOR_VERSION:=" - AutoHotkey v"A_AhkVersion
	MY_AHK = %MY_AHK%%AHK_TERMINATOR_VERSION%

	MY_AHK:=A_ScriptName 


	WinGet, Style, Style, %MY_AHK%
	If (Style & 0x10000000)  ; 0x10000000 is WS_VISIBLE
	{      ; MsgBox, My Window is Visible
		
	}
	Else  
	{
	SET_SHOW=
	; MsgBox, My Window is Hidden	
	}	

	WinGet MMX, MinMax, %MY_AHK%

	If MMX=0
	O_SET_SHOW_2=
	If MMX=1
	O_SET_SHOW_2=
	
	SET_SHOW_2=%MMX%
	
	IF O_SET_SHOW_2<>SET_SHOW_2
	IfEqual MMX,-1               ; -1 = Min __ IF MINIMIZED AND THEN HIDE
	{
		Gui, 1: Hide
	}
	O_SET_SHOW_2=%SET_SHOW_2%
	
	
RETURN


START:

	; ----
	; Command-line arguments - Rosetta Code
	; https://rosettacode.org/wiki/Command-line_arguments#AutoHotkey
	; ----
	; AutoHotkey
	; From the AutoHotkey documentation: "The script sees incoming parameters as the variables %1%, %2%, 
	; and so on. In addition, %0% contains the number of parameters passed (0 if none). "

	Loop %0% ; number of parameters
		Command_Params .= %A_Index% . A_Space
	IF %0%=0
		Command_Params=

	; ---------------------------------------------------------------
	; TEST DEBUGGE
	; RUN BY
	; Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
	; ---------------------------------------------------------------
	; Command_Params=/QUITE_COMMANDLINE_ARGS
	; ---------------------------------------------------------------

	; Each array must be initialized before use:
	FN_Array := []
	STATUS_Array := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array[ArrayCount]:="AU3_Spy.exe"            ; ---- AUTOHOTKEYS SPY PROGRAM
	ArrayCount += 1
	FN_Array[ArrayCount]:="Explorer.exe"
	ArrayCount += 1
	FN_Array[ArrayCount]:="Winamp.exe"             ; ---- MUSICAL MP3 AUDIO PROGRAM
	
	; ArrayCount += 1
	; FN_Array[ArrayCount]:="SystemSettings.exe"     
	; C:\Windows\ImmersiveControlPanel\SystemSettings.exe
	; FOUND TO BE IN SUSPENDED BUT TRIGGER RESUME COMES BACK TO SUSPENDED AGAIN BUT NOT ANY PROBLEM TO KILL PROCESS
	
	; ArrayCount += 1
	; FN_Array[ArrayCount]:="ShellExperienceHost.exe"     
	; C:\Windows\SystemApps\ShellExperienceHost_cw5n1h2txyewy\ShellExperienceHost.exe
	; FOUND TO BE IN SUSPENDED BUT TRIGGER RESUME COMES BACK TO SUSPENDED AGAIN BUT NOT ANY PROBLEM TO KILL PROCESS
	; AFE KIL PROCESS COMES BACK AGAIN BUT WITHOUT SUSPENDED ON
	; ----
	; ShellExperienceHost.exe Windows process - What is it?
	; https://www.file.net/process/shellexperiencehost.exe.html
	; ----


    Gui, Submit, NoHide
    WTSEnumProcesses(), LV_Delete(), count := 0
    loop % arrLIST.MaxIndex()
    {
        i := A_Index
		Loop % ArrayCount
		{
			; -------------------------------------------------------
			; USE THE ADDITIONAL "*" METHOD OR ELSE THINGS LIKE 
			; SYSTEMEXPLORER.EXE PROGRAM WOULD BE INCLUDED
			; InStr() CHECKER IS NOT CASE SENSITIVE
			; -------------------------------------------------------
			IF (InStr("*"arrLIST[i, "Process"], "*"FN_Array[A_Index]))
			{
				REASON_PROCESS_WAIT=
				IF IsProcessSuspended(arrLIST[i, "PID"])
				{
					LV_Add("", arrLIST[i, "PID"], arrLIST[i, "Process"], "Suspender" ), count++
					STATUS_Array[A_Index]:="INFO_RESULT_GIVEN_DONE"
					IF (count=1 and !SHOW_GUI)
					{
						SHOW_GUI:=TRUE
						Gui, Show, AutoSize
						SoundBeep , 1500 , 400
					}
				}
				ELSE
				{
					LV_Add("", arrLIST[i, "PID"], arrLIST[i, "Process"], "Not Suspended" )
					STATUS_Array[A_Index]:="INFO_RESULT_GIVEN_DONE"
					IF (count=1 and !SHOW_GUI and !Command_Params)
					{
						SHOW_GUI:=TRUE
						Gui, Show, AutoSize
						SoundBeep , 1500 , 400
					}
				}
			}
		}
	}
    GuiControl,, prccount, % count


IF count=0 
	GuiControl ,, Status, Not Any Problem to Worry About Your Loser
else
{
	IF count=1
		GuiControl ,, Status, Problem Founder
	IF count>1
		GuiControl ,, Status, Problems Found
}	
	 
IF count=0
	if Command_Params
		ExitApp

if !Command_Params
	Gui, Show, AutoSize

Loop % ArrayCount
{
		IF !STATUS_Array[A_Index]
			LV_Add("", "0", FN_Array[A_Index], "Not Runner" )
}
		
return

; FUNCTIONS =========================================================

WTSEnumProcesses()
{
    local tPtr := 0, pPtr := 0, nTTL := 0, LIST := ""
    ; DllCall("LoadLibrary", "Str", "Wtsapi32.dll")
    if !(DllCall("Wtsapi32\WTSEnumerateProcesses", "Ptr", 0, "Int", 0, "Int", 1, "PtrP", pPtr, "PtrP", nTTL))
        return "", DllCall("SetLastError", "Int", -1)

    tPtr := pPtr
    arrLIST := []
    loop % (nTTL)
    {
        arrLIST[A_Index, "PID"]     := NumGet(tPtr + 4, "UInt")    ; PID
        arrLIST[A_Index, "Process"] := StrGet(NumGet(tPtr + 8))    ; Process
        tPtr += (A_PtrSize = 4 ? 16 : 24)                          ; sizeof(WTS_PROCESS_INFO)
    }

    DllCall("Wtsapi32\WTSFreeMemory", "Ptr", pPtr)

    return arrLIST, DllCall("SetLastError", "UInt", nTTL)
}

; EXIT ==============================================================

GuiEscape:
GuiClose:
ExitApp

RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------

IsProcessSuspended(pid) {
    For thread in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Thread WHERE ProcessHandle = " pid)
        {
		If (thread.ThreadWaitReason != 5)
            Return False
		}
    Return True
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; USAGE EXAMPLE PUT THE TIMER ROUTINE TO PUT IN ANOTHER CODE AS LAUNCHER
;# ------------------------------------------------------------------

; -------------------------------------------------------------------
; INITIALIZING PART SET TIMER GOING - TO DELCARE AT BEGINNING OF CODE
; -------------------------------------------------------------------
SETTIMER TIMER_Check_Any_PID_Suspended_Warning, 10000 ; ---- 10 SECONDS ---- And Then 1 Hour

; -------------------------------------------------------------------
; ROUTINE TO ADD ANYWHERE IN CODE
; SET TO RUN ONLY ONCE IF ALREADY RUNNER
; -------------------------------------------------------------------
TIMER_Check_Any_PID_Suspended_Warning:
	SETTIMER TIMER_Check_Any_PID_Suspended_Warning, 3600000 ; ---- 10 SECONDS ---- And Then 1 Hour

	Element_1=C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 42-Check_Any_PID_Suspended_Warning.ahk

	SET_GO=FALSE
	IfExist, %Element_1%
		IF !WinExist(Element_1) 
			SET_GO=TRUE

	IF SET_GO=TRUE	
		{
			Run, "%Element_1%" /QUITE_COMMANDLINE_ARGS
		}
RETURN
 
;# ------------------------------------------------------------------
; USUAL END BLOCK OF CODE TO HELP EXIT ROUTINE
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000

if ScriptInstanceExist()
{
	Exitapp
}
return
; -------------------------------------------------------------------

; -------------------------------------------------------------------
ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % dhw
	return (match > 1)
	}
Return
; -------------------------------------------------------------------

; -------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
; -------------------------------------------------------------------

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))

; -------------------------------------------------------------------
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
; -------------------------------------------------------------------
; exit the app
; -------------------------------------------------------------------




; -------------------------------------------------------------------
; REFERENCE __ CODE SCARP BOOK
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; NOT USED MY ATTEMPT BETTER WAS FOUND
; ------------------------------------
TIMER_Check_Any_PID_Suspended_Fault:
Loop % ArrayCount
{
	VAR_PID_NAME := FN_Array[A_Index]

	Process, Exist, %VAR_PID_NAME%
	NewPID = %ErrorLevel%  
	If NewPID > 0 
	IF IsProcessSuspended(NewPID)
		MSGBOX % "IS SUSPENDED ---- " VAR_PID_NAME

}
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; INTERESTING isProcessSuspended RESULT ALTERNATIVE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; ----
; https://raw.githubusercontent.com/Drugoy/Autohotkey-scripts-.ahk/master/ScriptManager.ahk/MasterScript.ahk
; https://raw.githubusercontent.com/Drugoy/Autohotkey-scripts-.ahk/master/ScriptManager.ahk/MasterScript.ahk
; ----
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; Input: processID_N - ProcessID of any process.
; Output: boolean, where 'true' means the processe is suspended.
; Called by:
; Functions: memoryScan(), fillProcessesLV(), toggleSuspendProcess(), ProcessCreate_OnObjectReady().

isProcessSuspended_2(processID_N)
{	
	; 0 = Unknown, 1 = Other, 2 = Ready, 3 = Running, 4 = Blocked, 5 = Suspended Blocked, 6 = Suspended Ready.
	; http://msdn.microsoft.com/en-us/library/aa394372%28v=vs.85%29.aspx
	
	Global WMIQueries_O
	For thread In WMIQueries_O.ExecQuery("SELECT ThreadWaitReason FROM Win32_Thread WHERE ProcessHandle = " processID_N)
		If (thread.ThreadWaitReason == 5)
			Return 1	; Suspended.
	Return 0	; Not suspended.
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------




; -------------------------------------------------------------------
; REFERENCE WEB PAGES OPEN 18
; [ Thusday 11:59:00 Pm_18 October 2018 ]
; -------------------------------------------------------------------

; ----
; ahk suspend PID - Google Search
; https://www.google.co.uk/search?ei=YebIW8HHHoKxsQGHrZ3wCw&q=ahk+suspend+PID&oq=ahk+suspend+PID&gs_l=psy-ab.3...9003.10645.0.10881.9.7.0.0.0.0.291.860.0j3j1.4.0....0...1c.1.64.psy-ab..5.3.661...0j0i22i30k1j0i67k1j33i160k1.0.G0BBFX3DJD0
; --------
; Suspend a process OutSide AHK? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/119126-suspend-a-process-outside-ahk/
; --------
; Process suspend and unsuspend - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/78978-process-suspend-and-unsuspend/
; --------
; autohotkey - What is AU3_Spy.exe? Where can I find it? - Stack Overflow
; https://stackoverflow.com/questions/45552452/what-is-au3-spy-exe-where-can-i-find-it
; --------
; [Solved] Check if process is suspended - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/113862-solved-check-if-process-is-suspended/
; --------
; WinGetAll() - Get all windows Title/Class/PID/Process Name - Scripts and Functions - AutoHotkey Community
; https://autohotkey.com/board/topic/30323-wingetall-get-all-windows-titleclasspidprocess-name/
; --------
; Search Form - AutoHotkey Community
; https://autohotkey.com/board/index.php?app=core&module=search&section=search&do=search&fromsearch=1
; --------
; list of all processes - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/118907-list-of-all-processes/
; --------
; Change GUI font - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/64605-change-gui-font/
; --------
; AHK IS PROCESS SUSPEND DLL - Google Search
; https://www.google.co.uk/search?q=AHK+IS+PROCESS+SUSPEND+DLL&oq=AHK+IS+PROCESS+SUSPEND+DLL&aqs=chrome..69i57j69i64.10836j0j7&sourceid=chrome&ie=UTF-8
; --------
; scripting - Check if Script is suspended in AutoHotkey - Stack Overflow
; https://stackoverflow.com/questions/14492650/check-if-script-is-suspended-in-autohotkey
; --------
; AutoHotkey script to immediately kill an application when exceeded a set memory usage limit
; https://gist.github.com/cheeaun/130164
; --------
; Suspend a process OutSide AHK? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/119126-suspend-a-process-outside-ahk/
; --------
; Controller, Action or Item Not Found - CodeProject
; https://www.codeproject.com/threads/pausep.aspx
; --------
; NtSuspendProcess/NtSuspendProcess.au3 at master · jschicht/NtSuspendProcess
; https://github.com/jschicht/NtSuspendProcess/blob/master/NtSuspendProcess.au3
; --------
; Script written for basic: converting it to AHK_L - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/66559-script-written-for-basic-converting-it-to-ahk-l/
; --------
; How to make this Suspend/Resume process work with active window. - AutoHotkey Community
; https://autohotkey.com/boards/viewtopic.php?f=5&t=51663&p=227525
; --------
; https://raw.githubusercontent.com/Drugoy/Autohotkey-scripts-.ahk/master/ScriptManager.ahk/MasterScript.ahk
; https://raw.githubusercontent.com/Drugoy/Autohotkey-scripts-.ahk/master/ScriptManager.ahk/MasterScript.ahk
; ----

; -------------------------------------------------------------------
; REFERENCE WEB PAGES OPEN 7
; [ Friday 13:01:40 Pm_19 October 2018 ]
; -------------------------------------------------------------------
; ----
; command line - Autohotkey / AHK - Param %1% not accessible within a function - Stack Overflow
; https://stackoverflow.com/questions/49905729/autohotkey-ahk-param-1-not-accessible-within-a-function
; --------
; IsProcessSuspended(arrLIST - Google Search
; https://www.google.co.uk/search?num=50&safe=active&ei=aLHJW_2ZD8SQ5gKk9YzoCA&q=IsProcessSuspended%28arrLIST&oq=IsProcessSuspended%28arrLIST&gs_l=psy-ab.3...4053.4606.0.4993.4.4.0.0.0.0.148.465.1j3.4.0....0...1c.1.64.psy-ab..0.0.0....0.miN3IoeW0NM
; --------
; Matthew-Lancaster/Autokey -- 42-Check_Any_PID_Suspended_Warning.ahk at master · Matthew-Lancaster/Matthew-Lancaster
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2042-Check_Any_PID_Suspended_Warning.ahk
; --------
; Using IF to evaluate command line arguments - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/67927-using-if-to-evaluate-command-line-arguments/
; --------
; Scripts - Definition & Usage | AutoHotkey
; https://autohotkey.com/docs/Scripts.htm
; --------
; Command-line arguments - Rosetta Code
; https://rosettacode.org/wiki/Command-line_arguments#AutoHotkey
; --------
; multiple gui buttons?? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/3697-multiple-gui-buttons/
; ----