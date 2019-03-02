;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 42-Check_Any_PID_Suspended_Warning.ahk
;# __ 
;# __ Autokey -- 42-Check_Any_PID_Suspended_Warning.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;# __ DATE BEGIN
;# __ Thu 18-Oct-2018 21:10:00
;# __ 
;  =============================================================

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

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

; -------------------------------------------------------------------
#SingleInstance force
; -------------------------------------------------------------------



;# ------------------------------------------------------------------
; ---- LOCATION ONLINE
; -------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 42-Check_Any_PID_Suspended_Warning.ahk at master · Matthew-Lancaster/Matthew-Lancaster
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2042-Check_Any_PID_Suspended_Warning.ahk
; ----
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; DESCRIPTION
;# ------------------------------------------------------------------
;# ------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; ----
; CODER __ KNOCK THIS UP TO CHECK SUSPENDED PROCESS FAULTY
; HARD TO BECAUSE TRY AND FIND A DLL VERSION OF WMI
; CAN GET LIST SCRIPT OF PROCESS INSTANTLY BUT WMI TO CHECK IS SUSPENDED 
; NOT THAT GOOD FOR SPEED __ BIG SEARCH ON LINE LOTS OF READER
; -------------------------------------------------------------------
; WAS SUCH A LONG SEARCHER LOOKING FOR DLL METHOD OVER AGAINST LESS QUICKER WMI
; IT GET ALL PROCESSES SPEED OF LIGHT BUT CHECK EACH ONE OF WATCHED SELECTION 
; TAKE TIME __ EVERY TIME SOMEBODY WANTED THE DLL METHOD SOMEBODY ELSE BUTTS 
; IN WITH WMI METHOD AGAIN
; AND ALL TO TECHNICAL CLEVER TO GET OPEN PROCESS AND THEN FIND THE 
; COMMAND TO CHECK SUSPENDED _ I WISH SOME PEOPLE WOULD BE ABLE AND 
; CHECKED MANY CODE OVER
; SEEM POINTLESS TO PRAISE THE DLL AND THEN NOT HAVE 
; A SUSPEND CHECKER IN DLL ALSO
; MORE SEARCHING FOR CODE THAN GET ON WITH JOB AND LEAVE IT LIKE IT IS
; 9 PM TO 12 MIDNIGHT
; -------------------------------------------------------------------
; IF I WANTER TO CHECK ALL PROCESS FOR A COUNT OVER CERTAIN AMOUNT THAT 
; SUSPENDER DLL QUICKER VS WMI LESS QUICKER WOULD BE A PROBLEM
; -------------------------------------------------------------------
; FROM   Thu 18-Oct-2018 21:07:43
; TO     Thu 18-Oct-2018 23:58:00
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; SESSION 002
; -------------------------------------------------------------------
; ----
; CODER __ NEXT DAY AFTER WAKE UP AND LITTLE MINOR CODER
; PUT COMMAND LINE A_ARGS IN SO ABLE TO RUN IN QUITE MODE IF SPECIFY
; AND NORMAL RUN GIVE RESULT IF ANY OR SHOW EMPTY
; ADD MORE CODE TO GIVE RESULT BETTER FOR GUI RUN WITH OR WITHOUT A_ARGS
; READ SEARCH FURTHER DOWN FOR USAGE EXAMPLE
; -------------------------------------------------------------------
; FROM   Fri 19-Oct-2018 11:24:02
; TO     Fri 19-Oct-2018 13:05:00
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
; SoundBeep , 1500 , 400
; -------------------------------------------------------------------

; GLOBAL SETTINGS ===================================================

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

; GUI ===============================================================

Gui, Margin, 5, 5
gui, font, s14 ; , Arial ; , Calibri  
; Gui, Add, Edit, xm ym w200 hWndhSearch vprcsearch
; DllCall("user32.dll\SendMessage", "Ptr", hSearch, "UInt", 0x1501, "Ptr", 1, "Str", "Prozess Name here", "Ptr")
Gui, Add, Button, y+5 w540 gSTART_RERUN, Script of Watched Processes Suspended Warning
Gui, add, Button, y+5 W540 gStatus, Status
Gui, Add, ListView, xm y+5 w540 h340, PID|Name|HardCore
; Gui, Add, Edit, xm y+5 w60 0x800 vprccount
LV_ModifyCol(1, 100)
LV_ModifyCol(2, 240)
LV_ModifyCol(3, 200)

SETTIMER MINIMIZE_GUI,4000


GOSUB START

RETURN

; SCRIPT ============================================================
START_RERUN:
	Soundbeep 1000,100
	GOSUB START
RETURN

Status:
RETURN

MINIMIZE_GUI:

DetectHiddenWindows, on
ThisScriptsHWND := WinExist("Ahk_PID " DllCall("GetCurrentProcessId"))

WinMinimize ahk_id %ThisScriptsHWND%
SETTIMER MINIMIZE_GUI,OFF
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
	ArrayCount += 1
	FN_Array[ArrayCount]:="GoodSync-v10.exe"       ; ---- NOW GOODSYNC GOING ANNOYING CAN'T SEE MON 03 DEC 2018
	
	
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
		GuiControl ,, Status, Problem Set Found
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



MenuHandler:
	; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	if A_ThisMenuItem=Terminate Script
		Process,Close,% DllCall("GetCurrentProcessId")
	
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
		; Process,Close, AutoHotKey.exe
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
		; C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\BAT_03_PROCESS_KILLER.BAT
		; ORIGINAL AT HERE LOCATION 
		; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
		; AND MOVED HERE MAYBE 
		; -------------------------------------------------------------------
		; MOST LIKELY TRY AND KEEP IN SYNC LATER
		; EXCEPT THE AUTO GENERATOR
		; -------------------------------------------------------------------
}
return



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