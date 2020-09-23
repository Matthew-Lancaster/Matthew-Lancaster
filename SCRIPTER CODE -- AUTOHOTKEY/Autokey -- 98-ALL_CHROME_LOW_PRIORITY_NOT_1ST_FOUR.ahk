;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 94-ALL_LOWER_THAN_NORMAL_PROCCES_PRIORITY_RESTORE.ahk
;# __ 
;# __ Autokey -- 98-ALL_CHROME_LOW_PRIORITY_NOT_1ST_FOUR.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; NICE IDEA ONE
; IT MODIFY PROCESS PRIORITY TO ALL CHROME NOT THE FIRST FOUR
; WHEN LOOK PROCESS 1ST FOUR UN-SORT 
; FIRST IS WINDOW
; NEXT LOOK COMMAND-LINE
; 2 TO 3 ARE IMPORT FEATURE
; REST SOMETHING CALL RENDER
; BEEN MANY SOLUTION TO PUT TAB ASLEEP IDLE NOT USER
; GUESS ANOTHER WAY
; EACH ONE IS SOMETHING FOR EXTENSION
; -------------------------------------------------------------------
; UN-REM GUI SHOW LINE TO DISPLAY INFO
; -------------------------------------------------------------------
; HERE IS RUN EVERY HOUR FROM
; Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; ROUTINE 
; AUTOKEY__98_ALL_CHROME_LOW_PRIORITY_NOT_1ST_FOUR_AHK
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; Sun 20-Sep-2020 03:30:00
; Sun 20-Sep-2020 04:32:00 -- 1 HOUR 2 MINUTE
; -------------------------------------------------------------------

; ------------------------------------------------------------------
; Location SOURCE
;---------------------------------------------------------------------
; ---------------------
; Matthew-Lancaster/Autokey -- 98-ALL_CHROME_LOW_PRIORITY_NOT_1ST_FOUR.ahk
; ---------------------
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOHOTKEY/Autokey%20--%2098-ALL_CHROME_LOW_PRIORITY_NOT_1ST_FOUR.ahk
; ---------------------
; HTTPS://TINYURL.COM
; ---------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; ADDITIONAL CREDIT TO 
; -------------------------------------------------------------------
; PhiLho 
; Moderator 
; 6850 POST COUNT
; -------------------------------------------------------------------
; AHK Functions :: InCache() - Cache List of Recent Items - Page 3 - Scripts and Functions - AutoHotkey Community 
; -------------------------------------------------------------------
; https://autohotkey.com/board/topic/7984-ahk-functions-incache-cache-list-of-recent-items/://autohotkey.com/board/topic/7984-ahk-functions-incache-cache-list-of-recent-items/page-3?&#entry75675
; HTTP://TINYURL.COM/WK8RRGC
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; AND CREDIT TO **** AUTOHOTKEYS **** HELP PAGE -- SEARCH WORD -- Process
; -------------------------------------------------------------------
; WHICH I FIND TO MODIFY AND BRING OUT THE PID IN TEXT STRING OPERATION
; -------------------------------------------------------------------
; Process - Syntax & Usage | AutoHotkey 
; https://www.autohotkey.com/docs/commands/Process.htm
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; WITH MAIN MY CREDIT OPERATION
; A MASSIVE EFFORT HARD WORK FIND READY MADE LEARN SCRIPT 
; LIKE HERE
; -------------------------------------------------------------------
; Sat 29-Feb-2020 13:56:46
; Sat 29-Feb-2020 15:10:00 -- 1 HOUR 13 MINUTE
; Text Size  16303 -- CRC32 53AC75E0
; -------------------------------------------------------------------
; -------------------------------------------------------------------

#Warn
#NoEnv
#SingleInstance Force

; -------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
; OnExit("ExitFunc")

; Register an object to be called on exit:
; OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

DetectHiddenWindows, on

; -------------------------------------------------------------------
; CREDIT TO **** AUTOHOTKEYS **** HELP PAGE -- SEARCH WORD -- Process
; -------------------------------------------------------------------
; WHICH I FIND TO MODIFY AND BRING OUT THE PID IN TEXT STRING OPERATION
; -------------------------------------------------------------------
; Process - Syntax & Usage | AutoHotkey 
; https://www.autohotkey.com/docs/commands/Process.htm
; -------------------------------------------------------------------
t=
luid=
l=

d := "`n"  ; string separator
s := 4096  ; size of buffers and arrays (4 KB)

Process, Exist  ; Sets ErrorLevel to the PID of this running script.
; Get the handle of this script with PROCESS_QUERY_INFORMATION (0x0400):
h := DllCall("OpenProcess", "UInt", 0x0400, "Int", false, "UInt", ErrorLevel, "Ptr")
; Open an adjustable access token with this process (TOKEN_ADJUST_PRIVILEGES = 32):
DllCall("Advapi32.dll\OpenProcessToken", "Ptr", h, "UInt", 32, "PtrP", t)
VarSetCapacity(ti, 16, 0)  ; structure of privileges
NumPut(1, ti, 0, "UInt")  ; one entry in the privileges array...
; Retrieves the locally unique identifier of the debug privilege:
DllCall("Advapi32.dll\LookupPrivilegeValue", "Ptr", 0, "Str", "SeDebugPrivilege", "Int64P", luid)
NumPut(luid, ti, 4, "Int64")
NumPut(2, ti, 12, "UInt")  ; Enable this privilege: SE_PRIVILEGE_ENABLED = 2
; Update the privileges of this process with the new access token:
r := DllCall("Advapi32.dll\AdjustTokenPrivileges", "Ptr", t, "Int", false, "Ptr", &ti, "UInt", 0, "Ptr", 0, "Ptr", 0)
DllCall("CloseHandle", "Ptr", t)  ; Close this access token handle to save memory.
DllCall("CloseHandle", "Ptr", h)  ; Close this process handle to save memory.

hModule := DllCall("LoadLibrary", "Str", "Psapi.dll")  ; Increase performance by preloading the library.
s := VarSetCapacity(a, s)  ; An array that receives the list of process identifiers:
c := 0  ; counter for process idendifiers
DllCall("Psapi.dll\EnumProcesses", "Ptr", &a, "UInt", s, "UIntP", r)
Loop, % r // 4  ; Parse array for identifiers as DWORDs (32 bits):
{
    id := NumGet(a, A_Index * 4, "UInt")
    ; Open process with: PROCESS_VM_READ (0x0010) | PROCESS_QUERY_INFORMATION (0x0400)
    h := DllCall("OpenProcess", "UInt", 0x0010 | 0x0400, "Int", false, "UInt", id, "Ptr")
	if !h
        continue
    VarSetCapacity(n, s, 0)  ; a buffer that receives the base name of the module:
    e := DllCall("Psapi.dll\GetModuleBaseName", "Ptr", h, "Ptr", 0, "Str", n, "UInt", A_IsUnicode ? s//2 : s)
    if !e    ; fall-back method for 64-bit processes when in 32-bit mode:
        if e := DllCall("Psapi.dll\GetProcessImageFileName", "Ptr", h, "Str", n, "UInt", A_IsUnicode ? s//2 : s)
            SplitPath n, n
    DllCall("CloseHandle", "Ptr", h)  ; close process handle to save memory
    if (n && e)  ; if image is not null add to list:
        l .= n . " ---- " . id . d, c++ 
}
DllCall("FreeLibrary", "Ptr", hModule)  ; Unload the library to free memory.
; Sort, l, C  ; Uncomment this line to sort the list alphabetically.
; -------------------------------------------------------------------
; %c% Processes __ COUNT OF PROCESS
; MSGBOX, %c% Processes , %l%
; MSGBOX, %l%
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; DO NOTHING WITH THE FILTER 
; NOT WORTH BRING HERE PROCESS UP A BIT LEAVE THEM LOW
; ADD YOUR OWN __.EXE
; -------------------------------------------------------------------
FILTER=
FILTER=%FILTER%--GoogleCrashHandler64.exe`n
FILTER=%FILTER%--GoogleCrashHandler.exe`n
FILTER=%FILTER%--GoogleUpdate.exe`n
FILTER=%FILTER%--SearchFilterHost.exe`n
FILTER=%FILTER%--SearchProtocolHost.exe`n
FILTER=%FILTER%--`n
FILTER=%FILTER%--`n
FILTER=%FILTER%--`n
FILTER=%FILTER%--`n
FILTER=%FILTER%--`n
FILTER=%FILTER%--`n
FILTER=%FILTER%--`n
FILTER=%FILTER%--`n
FILTER=%FILTER%--`n
; -------------------------------------------------------------------
COUNTER_V=1
VAR_INFO=
PRIORITY_SET=
; -------------------------------------------------------------------
LOOP, parse, l, `n
{
	PIDV:=substr(A_LoopField, INSTR(A_LoopField,"----")+5)
	PIDNAME:=substr(A_LoopField,1, INSTR(A_LoopField,"----")-2)
	IF INSTR(FILTER,"--" . PIDNAME . "`n")=0
	IF PIDNAME=chrome.exe
	IF PIDV
	{
		IF GetPriority(PIDV)=1
			PRIORITY_SET=_LOW_______
		IF GetPriority(PIDV)=2
			PRIORITY_SET=_B_NORMAL_
		IF GetPriority(PIDV)=3
			PRIORITY_SET=_NORMAL___
		IF GetPriority(PIDV)=4
			PRIORITY_SET=_A_NORMAL_
			
		; -----------------------------------------------------------
		; THE CHROME.EXE ARRIVE AS SET THE FIRST 4
		; COMMAND LINE HAVE USER
		; AND REST TALK RENDER AT COMMAND LINE AMONG OTHER
		; -----------------------------------------------------------
		; Sun 20-Sep-2020 03:52:00
		; -----------------------------------------------------------
		
		Pack := "000000"
		PIDV_PAD = % (SubStr(Pack, 1, StrLen(Pack) - StrLen(PIDV)) . PIDV)
		Pack := "0000"
		CNTV_PAD = % (SubStr(Pack, 1, StrLen(Pack) - StrLen(COUNTER_V)) . COUNTER_V)
		SET_PID_PRIORITY=_NOT_MOVE__
		IF COUNTER_V>4
		SET_PID_PRIORITY=_LOW________

		VAR_INFO = % VAR_INFO " " CNTV_PAD " _ PRIOR FIND PRIORITY _ " GetPriority(PIDV) " " PRIORITY_SET "  PID " SET_PID_PRIORITY " PID SET " PIDV_PAD  " EXE " PIDNAME "`n"
			
		IF COUNTER_V>4
			Process, Priority, %PIDV%, Low
		COUNTER_V+=1
	}
}

; -------------------------------------------------------------------
DASH_STRING=------------------------------------------------------------------------------------------
VAR_INFO= % "  ALL LOWER PROCCES PRIORITY FOR CHROME EXCEPT 1ST TWO" "`n" DASH_STRING "`n" VAR_INFO DASH_STRING
; -------------------------------------------------------------------
; GUI ROUTINE HAS INFO ABOUT TIME OUT COUNTER
; -------------------------------------------------------------------
COUNTDOWN_TIME:=40
; -------------------------------------------------------------------
COUNTDOWN_TEXT_01=Second To EXIT Or ESC Key
COUNTDOWN_TEXT_02=%COUNTDOWN_TIME%  %COUNTDOWN_TEXT_01%

; DISPLAY FIND
; -------------------------------------------------------------------
Gui, -caption +toolwindow +AlwaysOnTop +Disabled -SysMenu +Owner  ; +Owner avoid taskbar button.
Gui, Color, White
Gui, Font, s11 bold, Arial
DASH_STRING=----------------------------------------------------
Gui, Add, Text, x10 y10 , % VAR_INFO
Gui, Add, Text, vCOUNTDOWN_TIME, %COUNTDOWN_TEXT_02%
; Gui, Show, NoActivate, Title of Window  ; NoActivate avoids deactivating the currently active window.
; Gui, Show
; -------------------------------------------------------------------
SETTIMER TIMER_EXIT_ROUTINE,1000
; -------------------------------------------------------------------

RETURN


; -------------------------------------------------------------------
; OnMessage(0x201, "WM_LBUTTONDOWN")
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; DONE LETS GET OUT OF HERE
; -------------------------------------------------------------------

; -------------------------------------------------------------------
RETURN
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; GUI TIMER CREDIT
; -------------------------------------------------------------------
; SKAN
; -------------------------------------------------------------------
; Administrator
; 9115 post
; Last active
; Joined: 26 Dec 2005
; -------------------------------------------------------------------
; Countdown in a GUI - Ask for Help - AutoHotkey Community 
; https://autohotkey.com/board/topic/32653-countdown-in-a-gui/
; -------------------------------------------------------------------
TIMER_EXIT_ROUTINE:
COUNTDOWN_TIME -= 1
COUNTDOWN_TEXT_02=%COUNTDOWN_TIME%  %COUNTDOWN_TEXT_01%
GuiControl,,COUNTDOWN_TIME,%COUNTDOWN_TEXT_02%
	IF COUNTDOWN_TIME<1 
	{
		Gui, Destroy
		EXITAPP
	}
	IF GetKeyState("LButton","P")
	{
		Gui, Destroy
		EXITAPP
	}
	IF GetKeyState("RButton","P")
	{
		Gui, Destroy
		EXITAPP
	}
RETURN
; -------------------------------------------------------------------

; -------------------------------------------------------------------
ESC::EXITAPP
; -------------------------------------------------------------------

; -------------------------------------------------------------------
Clicked_it:
	EXITAPP
RETURN
; -------------------------------------------------------------------


; -------------------------------------------------------------------
GetPriority(process="") {
	Process, Exist, %process%
	PID := ErrorLevel
	IfLessOrEqual, PID, 0, Return, "Error!"

	hProcess := DllCall("OpenProcess", Int,1024, Int,0, Int,PID)
	Priority := DllCall("GetPriorityClass", Int,hProcess)
	DllCall("CloseHandle", Int,hProcess)

	; IfEqual, Priority, 64   , Return, "Low"
	; IfEqual, Priority, 16384, Return, "BelowNormal"
	; IfEqual, Priority, 32   , Return, "Normal"
	; IfEqual, Priority, 32768, Return, "AboveNormal"
	; IfEqual, Priority, 128  , Return, "High"
	; IfEqual, Priority, 256  , Return, "Realtime"

	; ROUTINE CHECK FOR NUMERIC BELOW VALUE
	; ------------------------------------------------------------------
	IfEqual, Priority, 64   , Return, "1"
	IfEqual, Priority, 16384, Return, "2"
	IfEqual, Priority, 32   , Return, "3"
	IfEqual, Priority, 32768, Return, "4"
	IfEqual, Priority, 128  , Return, "5"
	IfEqual, Priority, 256  , Return, "8"
Return "" 
}
; -------------------------------------------------------------------
; Usage : GetPriority(processName)
; -------------------------------------------------------------------
; GetPriority() ascertains the priority level for an existing process and returns one of the following strings:
; -------------------------------------------------------------------
; Low
; BelowNormal
; Normal
; AboveNormal
; High
; RealtimeNote: If no parameter is passed, the script's own priority level is retrieved.
; -------------------------------------------------------------------
; MSDN Reference : GetPriorityClass() , OpenProcess() , CloseHandle() , Process Security and Access Rights
; -------------------------------------------------------------------
; Example along with the GetPriority() Function.
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; WIN_10 LOAD AS CALC.EXE AND THEN PROCESS NAME CALL Calculator.exe
; WHAT A PAIN FOR A DEMO
; Fri 28-Feb-2020 20:53:27 -- MATT.LAN@
; -------------------------------------------------------------------
; Run, Calc.exe
; WinWait, Calculator
; Process, Exist, Calculator.exe
; PIDV := ErrorLevel

; Process, Priority, %PIDV%, Low
; MsgBox, 64, Process Priority [Calculator.exe], % GetPriority("Calculator.exe")

; Process, Priority, %PIDV%, BelowNormal
; MsgBox, 64, Process Priority [Calculator.exe], % GetPriority("Calculator.exe")

; Process, Priority, %PIDV%, AboveNormal
; MsgBox, 64, Process Priority [Calculator.exe], % GetPriority("Calculator.exe")

; Process, Priority, %PIDV%, High
; MsgBox, 64, Process Priority [Calculator.exe], % GetPriority("Calculator.exe")

; Process, Priority, %PIDV%, Realtime
; MsgBox, 64, Process Priority [Calculator.exe], % GetPriority("Calculator.exe")

; Process, Priority, %PIDV%, Normal
; MsgBox, 64, Process Priority [Calculator.exe], % GetPriority("Calculator.exe")
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; NOT USER
; TEST CODE NOT FULL RESULT ONLY FINDER WINDOW TITLE
; -------------------------------------------------------------------
; VAR_PID2=
; WinGet, all, list ;get all hwnd
; Loop, %all%
; {
	; WinGet, PIDV, PID, % "ahk_id " all%A_Index%
	; IF PIDV
	; IF INSTR(VAR_PID2,PIDV . "`n")=0
	; {
		; VAR_PID2=%VAR_PID2%%PIDV%`n
		; IF GetPriority(PIDV)<3 
		; {
			; WinGet, PID_ProcessName, ProcessName, % "ahk_id " all%A_Index%
			; MSGBOX % GetPriority(PIDV) "--" PID_ProcessName
			; Process, Priority, %PIDV%, Normal
		; }
	; }
; }
; -------------------------------------------------------------------

; WM_LBUTTONDOWN(wParam, lParam)
; {
    ; X := lParam & 0xFFFF
    ; Y := lParam >> 16
    ; if A_GuiControl
        ; Control := "`n(in control " . A_GuiControl . ")"
    ; MSGBOX You left-clicked in Gui window #%A_Gui% at client coordinates %X%x%Y%.%Control%
; }

