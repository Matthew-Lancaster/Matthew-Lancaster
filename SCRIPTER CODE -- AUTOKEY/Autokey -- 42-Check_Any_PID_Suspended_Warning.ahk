;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 42-Check_Any_PID_Suspended_Warning.ahk
;# __ 
;# __ Autokey -- 42-Check_Any_PID_Suspended_Warning.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;# __ DATE BEGIN
;# __ Thu 18-Oct-2018 21:07:43
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
; DESCRIPTION
;# ------------------------------------------------------------------


;# ------------------------------------------------------------------
; ---- LOCATE ONLINE
; -------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 41-Minimize Chrome Close & Close RButton.ahk at master · Matthew-Lancaster/Matthew-Lancaster
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2041-Minimize%20Chrome%20Close%20%26%20Close%20RButton.ahk
; ----
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; ----
; CODER __ KNOCK THIS UP CHECK SUSPENDED PROCESS FAULTY
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
; TO     Fri 19-Oct-2018 11:30:00
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

; GLOBAL SETTINGS ===================================================

; DUPLICATE ENTRY TURNED OFF
; -------------------------------------------------------------------
; #Warn
; #NoEnv
; #SingleInstance Force

; GLOBAL REASON_PROCESS_WAIT
; REASON_PROCESS_WAIT=


; GUI ===============================================================

Gui, Margin, 5, 5
gui, font, s14 ; , Calibri  
; Gui, Add, Edit, xm ym w200 hWndhSearch vprcsearch
; DllCall("user32.dll\SendMessage", "Ptr", hSearch, "UInt", 0x1501, "Ptr", 1, "Str", "Prozess Name here", "Ptr")
; DllCall("user32.dll\SendMessage", , , , , , , , , "Ptr")
Gui, Add, Button, x+5 yp-1 w480 gSTART_2, Script of Watched Processes Suspended Warning
Gui, Add, ListView, xm y+5 w480 h300, PID|Name|HardCore
; Gui, Add, Edit, xm y+5 w60 0x800 vprccount
LV_ModifyCol(1, 100)
LV_ModifyCol(2, 180)
LV_ModifyCol(3, 150)

GOSUB START

RETURN

; SCRIPT ============================================================
START_2:
	Soundbeep 1000,100
	GOSUB START
RETURN

START:

	; Each array must be initialized before use:
	FN_Array := []
	STATUS_Array := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array[ArrayCount]:="AU3_Spy.exe"    ; ---- AUTOHOTKEYS SPY PROGRAM
	ArrayCount += 1
	FN_Array[ArrayCount]:="Explorer.exe"
	ArrayCount += 1
	FN_Array[ArrayCount]:="Winamp.exe"     ; ---- MUSICAL MP3 AUDIO PROGRAM

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
					LV_Add("", arrLIST[i, "PID"], arrLIST[i, "Process"], "Not Suspended" ), count++
					STATUS_Array[A_Index]:="INFO_RESULT_GIVEN_DONE"
					IF (count=1 and !SHOW_GUI)
					{
						if !A_Args
						{
							SHOW_GUI:=TRUE
							Gui, Show, AutoSize
							SoundBeep , 1500 , 400
						}
						
					}
				
				}
			}
		}
	}
    GuiControl,, prccount, % count

IF count=0 
	if A_Args
		ExitApp

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



;# ------------------------------------------------------------------
; USAGE EXAMPLE PUT THE TIMER ROUTINE IN ANOTHER CODE AS LAUNCHER
;# ------------------------------------------------------------------

; -------------------------------------------------------------------
; INITIALIZING PART SET TIMER GOING
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
			Run, "%Element_1%"
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
; REFERENCE PAGES OPEN
; -------------------------------------------------------------------
